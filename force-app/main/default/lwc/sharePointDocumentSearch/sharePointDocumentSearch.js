import { LightningElement } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { reduceErrors } from 'c/ldsUtils';
import searchDocuments from '@salesforce/apex/SharePointSearchController.searchDocuments';
import getDownloadUrl from '@salesforce/apex/SharePointSearchController.getDownloadUrl';

export default class SharePointDocumentSearch extends LightningElement {
    documentName = '';
    author = '';

    results = [];
    error;
    isLoading = false;
    hasSearched = false;

    handleDocumentNameChange(event) {
        this.documentName = event.target.value;
    }

    handleAuthorChange(event) {
        this.author = event.target.value;
    }

    handleClear() {
        this.documentName = '';
        this.author = '';
        this.results = [];
        this.error = undefined;
        this.hasSearched = false;
    }

    async handleSearch() {
        this.isLoading = true;
        this.hasSearched = true;
        this.error = undefined;
        try {
            const searchResults = await searchDocuments({
                documentName: this.documentName,
                author: this.author
            });
            this.results = searchResults.map((doc, index) => ({
                key: index,
                ...doc
            }));
        } catch (error) {
            this.error = error;
            this.results = [];
        } finally {
            this.isLoading = false;
        }
    }

    async handleDownloadClick(event) {
        const { driveId, itemId } = event.target.dataset;
        try {
            const downloadUrl = await getDownloadUrl({ driveId, itemId });
            window.open(downloadUrl, '_blank');
        } catch (error) {
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Unable to download document',
                    message: reduceErrors(error).join(', '),
                    variant: 'error'
                })
            );
        }
    }

    get hasResults() {
        return this.results.length > 0;
    }

    get showNoResults() {
        return (
            this.hasSearched &&
            !this.isLoading &&
            !this.error &&
            !this.hasResults
        );
    }
}
