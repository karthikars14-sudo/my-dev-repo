import { createElement } from '@lwc/engine-dom';
import SharePointDocumentSearch from 'c/sharePointDocumentSearch';
import { ShowToastEventName } from 'lightning/platformShowToastEvent';
import searchDocuments from '@salesforce/apex/SharePointSearchController.searchDocuments';
import getDownloadUrl from '@salesforce/apex/SharePointSearchController.getDownloadUrl';

// Mocking imperative Apex method calls
jest.mock(
    '@salesforce/apex/SharePointSearchController.searchDocuments',
    () => ({
        default: jest.fn()
    }),
    { virtual: true }
);
jest.mock(
    '@salesforce/apex/SharePointSearchController.getDownloadUrl',
    () => ({
        default: jest.fn()
    }),
    { virtual: true }
);

const SEARCH_RESULTS = [
    {
        title: 'testdocument.docx',
        documentType: 'Contract',
        author: 'Jane Doe',
        projectId: '12345',
        opportunityId: '00654321',
        fileUrl:
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx',
        lastModified: '2026-06-01T10:00:00Z',
        driveId: 'drive-1',
        itemId: 'item-1'
    }
];

const SEARCH_ERROR = {
    body: { message: 'An internal server error has occurred' },
    ok: false,
    status: 400,
    statusText: 'Bad Request'
};

const DOWNLOAD_ERROR = {
    body: { message: 'Unable to get a download link' },
    ok: false,
    status: 400,
    statusText: 'Bad Request'
};

describe('c-share-point-document-search', () => {
    let windowOpenSpy;

    beforeEach(() => {
        windowOpenSpy = jest.spyOn(window, 'open').mockImplementation(() => {});
    });

    afterEach(() => {
        while (document.body.firstChild) {
            document.body.removeChild(document.body.firstChild);
        }
        jest.clearAllMocks();
        windowOpenSpy.mockRestore();
    });

    // Helper function to wait until the microtask queue is empty. This is needed for promise
    // timing when calling imperative Apex.
    async function flushPromises() {
        return Promise.resolve();
    }

    function getFieldInput(element, label) {
        return Array.from(
            element.shadowRoot.querySelectorAll('lightning-input')
        ).find((input) => input.label === label);
    }

    async function searchAndGetElement() {
        searchDocuments.mockResolvedValue(SEARCH_RESULTS);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();

        await flushPromises();
        await flushPromises();

        return element;
    }

    it('renders all 5 search fields and a Clear and Search button', () => {
        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const labels = Array.from(
            element.shadowRoot.querySelectorAll('lightning-input')
        ).map((input) => input.label);
        expect(labels).toEqual([
            'Document Name',
            'Document Type',
            'Project ID',
            'Opportunity ID',
            'Author'
        ]);

        const buttonLabels = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).map((button) => button.label);
        expect(buttonLabels).toEqual(['Clear', 'Search']);
    });

    it('clears all field values and results when Clear is clicked', async () => {
        searchDocuments.mockResolvedValue(SEARCH_RESULTS);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const nameInput = getFieldInput(element, 'Document Name');
        nameInput.value = 'test';
        nameInput.dispatchEvent(new CustomEvent('change'));

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();
        await flushPromises();
        await flushPromises();

        expect(
            element.shadowRoot.querySelectorAll('table tbody tr').length
        ).toBe(1);

        const clearButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Clear');
        clearButton.click();
        await flushPromises();

        expect(getFieldInput(element, 'Document Name').value).toBe('');
        expect(element.shadowRoot.querySelector('table')).toBeNull();
    });

    it('triggers a search when Enter is pressed in a search field', async () => {
        searchDocuments.mockResolvedValue(SEARCH_RESULTS);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const nameInput = getFieldInput(element, 'Document Name');
        nameInput.value = 'test';
        nameInput.dispatchEvent(new CustomEvent('change'));
        nameInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter' }));

        await flushPromises();
        await flushPromises();

        expect(searchDocuments).toHaveBeenCalledTimes(1);
        expect(
            element.shadowRoot.querySelectorAll('table tbody tr').length
        ).toBe(1);
    });

    it('does not trigger a search when a non-Enter key is pressed', async () => {
        searchDocuments.mockResolvedValue(SEARCH_RESULTS);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const nameInput = getFieldInput(element, 'Document Name');
        nameInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'a' }));

        await flushPromises();

        expect(searchDocuments).not.toHaveBeenCalled();
    });

    it('calls Apex with all 5 search field values and renders a result row', async () => {
        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);
        searchDocuments.mockResolvedValue(SEARCH_RESULTS);

        getFieldInput(element, 'Document Name').value = 'test';
        getFieldInput(element, 'Document Name').dispatchEvent(
            new CustomEvent('change')
        );
        getFieldInput(element, 'Document Type').value = 'Contract';
        getFieldInput(element, 'Document Type').dispatchEvent(
            new CustomEvent('change')
        );
        getFieldInput(element, 'Project ID').value = '12345';
        getFieldInput(element, 'Project ID').dispatchEvent(
            new CustomEvent('change')
        );
        getFieldInput(element, 'Opportunity ID').value = '00654321';
        getFieldInput(element, 'Opportunity ID').dispatchEvent(
            new CustomEvent('change')
        );
        getFieldInput(element, 'Author').value = 'Jane Doe';
        getFieldInput(element, 'Author').dispatchEvent(
            new CustomEvent('change')
        );

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();

        await flushPromises();
        await flushPromises();

        expect(searchDocuments.mock.calls[0][0]).toEqual({
            documentName: 'test',
            documentType: 'Contract',
            projectId: '12345',
            opportunityId: '00654321',
            author: 'Jane Doe'
        });

        const rows = element.shadowRoot.querySelectorAll('table tbody tr');
        expect(rows.length).toBe(1);
        const downloadButton = rows[0].querySelector('lightning-button');
        expect(downloadButton.label).toBe('Download');
    });

    it('fetches the download URL on demand and opens it when Download is clicked', async () => {
        const element = await searchAndGetElement();
        getDownloadUrl.mockResolvedValue(
            'https://download.sharepoint.com/direct/testdocument.docx'
        );

        const downloadButton = element.shadowRoot.querySelector(
            'table tbody lightning-button'
        );
        downloadButton.click();

        await flushPromises();
        await flushPromises();

        expect(getDownloadUrl.mock.calls[0][0]).toEqual({
            driveId: 'drive-1',
            itemId: 'item-1'
        });
        expect(windowOpenSpy).toHaveBeenCalledWith(
            'https://download.sharepoint.com/direct/testdocument.docx',
            '_blank'
        );
    });

    it('shows an error toast when the download URL lookup fails', async () => {
        const element = await searchAndGetElement();
        getDownloadUrl.mockRejectedValue(DOWNLOAD_ERROR);

        const handler = jest.fn();
        element.addEventListener(ShowToastEventName, handler);

        const downloadButton = element.shadowRoot.querySelector(
            'table tbody lightning-button'
        );
        downloadButton.click();

        await flushPromises();
        await flushPromises();

        expect(handler).toHaveBeenCalled();
        expect(handler.mock.calls[0][0].detail.variant).toBe('error');
        expect(windowOpenSpy).not.toHaveBeenCalled();
    });

    it('shows a "no documents found" message when the search returns no results', async () => {
        searchDocuments.mockResolvedValue([]);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();

        await flushPromises();
        await flushPromises();

        expect(element.shadowRoot.textContent).toContain(
            'No documents found matching your search criteria.'
        );
    });

    it('renders the error panel when the Apex method returns an error', async () => {
        searchDocuments.mockRejectedValue(SEARCH_ERROR);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();

        await flushPromises();
        await flushPromises();

        const errorPanelEl = element.shadowRoot.querySelector('c-error-panel');
        expect(errorPanelEl).not.toBeNull();
    });

    it('is accessible when results are returned', async () => {
        const element = await searchAndGetElement();

        await expect(element).toBeAccessible();
    });

    it('is accessible when an error is returned', async () => {
        searchDocuments.mockRejectedValue(SEARCH_ERROR);

        const element = createElement('c-share-point-document-search', {
            is: SharePointDocumentSearch
        });
        document.body.appendChild(element);

        const searchButton = Array.from(
            element.shadowRoot.querySelectorAll('lightning-button')
        ).find((button) => button.label === 'Search');
        searchButton.click();

        await flushPromises();
        await flushPromises();

        await expect(element).toBeAccessible();
    });
});
