import * as React from "react";
import { Dialog, DialogFooter, DialogType, PrimaryButton, DefaultButton, Spinner, SpinnerSize } from "@fluentui/react";
import { getSP } from "../../../pnpjs-config";
import { ListViewCommandSetContext, RowAccessor } from "@microsoft/sp-listview-extensibility";
import { Logger } from "@pnp/logging";
import strings from "ApplicableMenuCommandSetStrings";
import { ListUrls } from "../../../constants";

const LOG_SOURCE: string = 'ApplicableMenuCommandSet';

export interface IWorkflowDialogContentProps {
    onClose: () => void;
    onConfirm: () => void;
    getSelection: () => readonly RowAccessor[] | undefined;
    getselectedId: () => string | undefined;
    fileLeafRefs: string[] | undefined;
    context: ListViewCommandSetContext;
}

interface IConfirmationPopupState {
    isDialogOpen: boolean;
    ConfirmationMessage: string;
    isLoading?: boolean;
    isError?: boolean;
    result?: { ref: string; rev: string; success: boolean }[];
    foundLibraries: { name: string; libraryName: string }[];
    isExistedItems: boolean;
    checkItems?: boolean;
    draftLibraryId?: string;
    applicableLibraryId?: string;
    prevLibraryId?: string;
    errorMessages: string[];
}

class ConfirmationPopup extends React.Component<IWorkflowDialogContentProps, IConfirmationPopupState> {
    constructor(props: IWorkflowDialogContentProps) {
        super(props);
        this.state = {
            isDialogOpen: true,
            ConfirmationMessage: strings.ConfirmationMessage,
            isExistedItems: false,
            foundLibraries: [],
            errorMessages: [],
        };
        this.onCancel = this.onCancel.bind(this);
    }

    public componentDidMount(): void {
        const { context } = this.props;
        const { errorMessages } = this.state;
        const sp = getSP();
        const [batch, execute] = sp.batched();

        Promise.all([
            batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.Draft}'`)(),
            batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.ApplicableDocuments}'`)(),
            batch.web.lists.select('Id').filter(`RootFolder/ServerRelativeUrl eq '${context.pageContext.web.serverRelativeUrl}/${ListUrls.PreviousVersions}'`)(),
        ])
            .then(([draftLib, applicableLib, prevLib]) => {
                if (!draftLib || draftLib.length === 0 || !applicableLib || applicableLib.length === 0 || !prevLib || prevLib.length === 0) {
                    this.setState({
                        errorMessages: [...errorMessages, strings.errorMessage],
                    });
                    console.error(new Error('Document libraries not found'));
                    Logger.error(new Error(`${LOG_SOURCE}: Document libraries not found`));
                } else {
                    this.setState({
                        draftLibraryId: draftLib[0].Id,
                        applicableLibraryId: applicableLib[0].Id,
                        prevLibraryId: prevLib[0].Id,
                    });
                }
            })
            .catch((error) => {
                this.setState({
                    errorMessages: [...errorMessages, strings.errorMessage],
                });
                console.error(error);
                Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
            });

        execute().catch((error) => {
            this.setState({
                errorMessages: [...errorMessages, strings.errorMessage],
            });
            console.error(error);
            Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
        });
    }

    // Hide the dialog
    hideDialog = () => {
        this.setState({ isDialogOpen: false });
        this.props.onClose();
    };

    // Handle Cancel button click
    private onCancel(): void {
        this.hideDialog();
    }

    // Search for files using SharePoint Search API
    private searchFiles = async (fileLeafRefs: string[]): Promise<{ name: string; libraryName: string }[]> => {
        const sp = getSP();
        const searchQuery = fileLeafRefs.map((fileLeafRef) => `FileLeafRef:${fileLeafRef}`).join(' OR ');

        try {
            const results = await sp.search({
                Querytext: searchQuery,
                RowLimit: 100,
                SelectProperties: ["FileLeafRef", "ListTitle"],
            });

            return results.PrimarySearchResults.map((result) => ({
                name: result.FileLeafRef,
                libraryName: result.ListTitle,
            }));
        } catch (error) {
            console.error("Error searching for files:", error);
            Logger.error(new Error(`${LOG_SOURCE}: ${error}`));
            return [];
        }
    };

    private _changeDocumentStatusToApplicable = async (): Promise<void> => {
        this.setState({ isLoading: true, isError: false });

        const { fileLeafRefs, getSelection, getselectedId } = this.props;
        const selectedRows = getSelection();

        if (!fileLeafRefs || fileLeafRefs.length === 0 || !selectedRows || selectedRows.length === 0) {
            this.setState({ isLoading: false, isError: true });
            return;
        }

        try {
            // Check for file existence using the Search API
            const foundLibraries = await this.searchFiles(fileLeafRefs);

            if (foundLibraries.length > 0) {
                // If files exist, show the details and stop further processing
                this.setState({ isExistedItems: true, foundLibraries, checkItems: true, isLoading: false });
                return;
            }

            // If no files exist, proceed to update the document status
            const sp = getSP();
            const [batch, execute] = sp.batched();
            const result: { ref: string; rev: string; success: boolean }[] = [];

            selectedRows.forEach((row) => {
                try {
                    const documentStatusField = row.getValueByName("DocumentStatus");
                    const selected = getselectedId();
                    if (documentStatusField === "Draft" && selected) {
                        batch.web.lists
                            .getById(selected)
                            .items.getById(row.getValueByName("ID"))
                            .update({ DocumentStatus: "Applicable" })
                            .then((response) => {
                                result.push({
                                    ref: row.getValueByName("ProjectReference"),
                                    rev: row.getValueByName("ProjectRevision"),
                                    success: true,
                                });
                            })
                            .catch((error) => {
                                console.error(error);
                                Logger.error(error);
                                result.push({
                                    ref: row.getValueByName("ProjectReference"),
                                    rev: row.getValueByName("ProjectRevision"),
                                    success: false,
                                });
                            });
                    }
                } catch (error) {
                    console.error("Error updating document status:", error);
                }
            });

            await execute();
            this.setState({ isLoading: false, result });
        } catch (error) {
            console.error("Error checking file existence:", error);
            this.setState({ isLoading: false, isError: true });
        }
    };

    render() {
        const { isDialogOpen, isLoading, isError, result, checkItems, foundLibraries } = this.state;
        let message = "";
        if (!isError && !isLoading && !result) {
            message = strings.ConfirmationMessage;
        }
        if (isError) {
            message = strings.errorMessage;
        }
        if (result) {
            message = strings.resultMessage
                .replace("{0}", result.filter((r) => r.success)?.length?.toString())
                .replace("{1}", result.length?.toString());
        }

        return (
            <>
                <Dialog
                    hidden={!isDialogOpen}
                    onDismiss={this.hideDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: strings.DialogTitle,
                        subText: `${checkItems ? strings.existedMessage : message}`,
                    }}
                >
                    {isLoading && <Spinner size={SpinnerSize.medium} />}

                    {checkItems && (
                        <div>
                            <span>Details</span>
                            <ul>
                                {foundLibraries.map((item, index) => (
                                    <li key={index}>
                                        {item.name} in {item.libraryName}
                                    </li>
                                ))}
                            </ul>
                        </div>
                    )}

                    <DialogFooter>
                        {(result || isError) && <PrimaryButton onClick={this.onCancel} text="OK" />}
                        {!isError && !isLoading && !result && !checkItems && (
                            <PrimaryButton onClick={this._changeDocumentStatusToApplicable.bind(this)} text="OK" />
                        )}
                        {!isError && !isLoading && !result && <DefaultButton onClick={this.onCancel} text="Cancel" />}
                    </DialogFooter>
                </Dialog>
            </>
        );
    }
}

export default ConfirmationPopup;
