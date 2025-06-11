import { LightningElement, api, track } from 'lwc';
import { ShowToastEvent } from "lightning/platformShowToastEvent";
import getFolderContentsFromRecord from '@salesforce/apex/SharePointFileBrowserController.getFolderContentsFromRecord';
import getFolderChildrenById from '@salesforce/apex/SharePointFileBrowserController.getFolderChildrenById';
import uploadFileToSharePoint from '@salesforce/apex/SharePointFileBrowserController.uploadFileToSharePoint';
import getSharePointPreviewUrl from '@salesforce/apex/SharePointFileBrowserController.getSharePointPreviewUrl';
import createFolderInSharePoint from '@salesforce/apex/SharePointFileBrowserController.createFolderInSharePoint';
import deleteSharepointFile from '@salesforce/apex/SharePointFileBrowserController.deleteSharepointFile';
import searchSharePoint from '@salesforce/apex/SharePointFileBrowserController.searchSharePoint';
import checkSharePointToken from '@salesforce/apex/SharePointOauthController.checkSharePointToken';
import initiateAuthFlow from '@salesforce/apex/SharePointFileBrowserController.initiateAuthFlow';
// TODO: Platform Events to signal success?
// Access on RFA object using Sharepoint link

export default class SharePointBrowser extends LightningElement {
    @api recordId;
    @api objectApiName;
    
    @track files = [];
    @track error = {};
    @track breadcrumbs = [];
    isCheckingAuth = true;
    isAuthenticated = false;
    showLoginButton = false;
    authPopup = null;


    isLoading = false;
    needsAuth = false;

    showUploadModal = false;
    showPreviewModal = false;
    showCreateFolderModal = false;
    showDeleteConfirmModal = false;

    deleteFileId;
    deleteFileName;
    currentItemId;
    currentDriveId;
    previewUrl = '';
    currentPath = '';
    downloadUrl = '';
    newFolderName = '';
    previewFileUrl = '';
    previewFileName = '';

    // Search params:
    showSearchInput = false;
    searchTerm = '';
    searchTimeout;
    

    connectedCallback() {
        // we want to know when the popup tells us it's done
        window.addEventListener('message', this.handleAuthMessage);

        // check if we already have a token
        checkSharePointToken()
        .then(hasToken => {
            this.isAuthenticated  = hasToken;
            this.isCheckingAuth   = false;

            if (hasToken) {
                this.loadRootFolder();
            }
        })
        .catch(err => {
            if (err?.body?.message?.includes('No token record found for this user')) {
                console.warn('User is not authorized with token:', err);
            } else {
                console.error('checkSharePointToken error', err);
            }
            this.isCheckingAuth = false;
            this.isAuthenticated = false;
        });
    }
        
    async handleLoginClick() {
        try {
            const authUrl = await initiateAuthFlow();
            window.open(
            authUrl,
            'SPAuthWindow',
            'width=600,height=700,menubar=no,toolbar=no,location=no,status=no'
            );
        } catch (err) {
            console.error('initiateAuthFlow error', err);
        }
    }

    get showLoginScreen() {
        return !this.isCheckingAuth && !this.isAuthenticated;
    }

    handleAuthMessage = (event) => {
    // (optional) check event.origin here…
    if (event.data === 'SP_AUTH_SUCCESS') {
        this.isAuthenticated = true;      // ← new!
        this.dispatchEvent(new ShowToastEvent({
        title: 'Success',
        message: 'You are signed in and may browser SharePoint files.',
        variant: 'success'
        }));
        this.loadRootFolder();
        window.removeEventListener('message', this.handleAuthMessage);
    }
    }


    // renderedCallback() {
    //     console.log('breadcrumbs:', JSON.stringify(this.breadcrumbs, null, '\t'));
    //     console.log(`current drive ID: ${this.currentDriveId} and item ${this.currentItemId}`)
    // }

    refreshData() {
        if (this.currentDriveId && this.currentItemId) {
            this.loadFolderById(this.currentDriveId, this.currentItemId);
        } else {
            this.loadRootFolder();
        }
    }

    async checkAuthStatus() {
        try {
            // Replace this with a real Apex call to check session validity
            const result = await checkSharePointSession(); // e.g., returns true/false
            this.needsAuth = !result;
            if (!this.needsAuth) {
                this.loadRootFolder();
            }
        } catch (error) {
            console.error('Auth check failed:', error);
            this.needsAuth = true;
        }
    }


    async loadRootFolder() {
        this.isLoading = true;
        try {
            const hasToken = await checkSharePointToken();
            if (!hasToken) {
                await this.redirectToLogin();
                return;
            }

            console.debug('Token check passed, loading files....');
            const raw = await getFolderContentsFromRecord({
                recordId: this.recordId,
                objectApiName: this.objectApiName
            });

            const data = JSON.parse(JSON.stringify(raw));
            this.currentPath = '';
            this.currentItemId = null;
            this.currentDriveId = null;
            this.setBreadcrumbs([]);
            this.files = this.formatFiles(data);

            // *Persist* the root folder driveId and itemId for future use
            if (data.length > 0) {
                const parentRef = data[0].parentReference;
                if (parentRef) {
                    this.currentDriveId = parentRef.driveId;
                    this.currentItemId = parentRef.id;
                }
            }
        } catch (error) {
            this.error = {
                message: error?.body?.message || error.message,
                isException: (error?.body?.message || error.message || '').includes('Missing SharePoint link')
            };
        } finally {
            this.isLoading = false;
        }
    }



    async loadFolderById(driveId, itemId) {
        this.isLoading = true;
        this.error = null;
        try {
            const data = await getFolderChildrenById({
                driveId,
                itemId
            });
            this.files = this.formatFiles(data);
        } catch (err) {
            this.error.message = error?.body?.message || error.message;
            this.error.isException = this.error.includes('Missing SharePoint link');
            this.files = [];
        } finally {
            this.isLoading = false;
        }
    }

    handleFolderClick(event) {
        const folderId = event.currentTarget.dataset.id;
        const driveId = event.currentTarget.dataset.driveid;
        const folderName = event.currentTarget.dataset.name;
        const nextPath = [...this.breadcrumbs.map(b => b.label), folderName];
        this.breadcrumbs = nextPath.map((label, i) => ({
            label,
            index: i,
            isLast: i === nextPath.length - 1,
            itemId: i === nextPath.length - 1 ? folderId : this.breadcrumbs[i]?.itemId,
            driveId: i === nextPath.length - 1 ? driveId : this.breadcrumbs[i]?.driveId
        }));

        console.log('breadcrumbs: ', JSON.stringify(this.breadcrumbs, null, '\t'))
        this.currentItemId = folderId;
        this.currentDriveId = driveId;
        this.loadFolderById(driveId, folderId);
    }

    handleBreadcrumbClick(event) {
        event.preventDefault();
        const index = parseInt(event.currentTarget.dataset.index, 10);
        const crumb = this.breadcrumbs[index];
        const newPath = this.breadcrumbs.slice(0, index + 1);

        this.breadcrumbs = newPath.map((b, i) => ({
            label: b.label,
            index: i,
            isLast: i === newPath.length - 1,
            itemId: b.itemId,
            driveId: b.driveId
        }));

        if (!crumb.itemId || !crumb.driveId) {
            this.loadRootFolder();
        } else {
            this.loadFolderById(crumb.driveId, crumb.itemId);
        }
    }

    handleBackClick() {
        if (this.breadcrumbs.length <= 1) {
            this.loadRootFolder();
        } else {
            const newBreadcrumbs = this.breadcrumbs.slice(0, -1);
            const lastCrumb = newBreadcrumbs[newBreadcrumbs.length - 1];

            this.breadcrumbs = newBreadcrumbs.map((b, i) => ({
                label: b.label,
                index: i,
                isLast: i === newBreadcrumbs.length - 1,
                itemId: b.itemId,
                driveId: b.driveId
            }));

            if (!lastCrumb.itemId || !lastCrumb.driveId) {
                this.loadRootFolder();
            } else {
                this.loadFolderById(lastCrumb.driveId, lastCrumb.itemId);
            }
        }
    }

    setBreadcrumbs(pathParts) {
        const lastIndex = pathParts.length - 1;
        this.breadcrumbs = pathParts.map((label, i) => ({
            label,
            index: i,
            isLast: i === lastIndex
        }));
    }

    handleUploadOpen() {
        this.showUploadModal = true;
    }

    handleUploadClosed() {
        this.showUploadModal = false;
    }

    /**
     * Button select upload 
     */
    handleUploadFinished(event) {
        this.isLoading = true;
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();

        reader.onloadend = async () => {
            const base64Data = reader.result.split(',')[1]; // Get just the base64 string
            try {
                await uploadFileToSharePoint({
                    base64Body: base64Data,
                    fileName: file.name,
                    driveId: this.currentDriveId,
                    itemId: this.currentItemId
                });
                this.showUploadModal = false;
                this.uploadSuccessMessage(file?.name);
                await this.loadFolderById(this.currentDriveId, this.currentItemId);
            } catch (error) {
                console.error('Upload failed:', error);
                this.error = {
                    message: error?.body?.message || error.message
                };
            } finally {
                this.isLoading = false;
            }
        };

        reader.readAsDataURL(file); // Triggers reader.onloadend, finishing upload
    }


    promptDeleteFile(event) {
        this.deleteFileId = event.currentTarget.dataset.id;
        this.deleteFileName = event.currentTarget.dataset.name;
        this.showDeleteConfirmModal = true;
    }

    cancelDelete() {
        this.deleteFileId = null;
        this.deleteFileName = '';
        this.showDeleteConfirmModal = false;
    }

    async confirmDelete() {
        try {
            await deleteSharepointFile({
                itemId: this.deleteFileId,
                driveId: this.currentDriveId
            });

            this.showDeleteConfirmModal = false;
            this.deleteFileId = null;
            this.deleteFileName = '';
            this.loadFolderById(this.currentDriveId, this.currentItemId);

            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Deleted',
                    message: 'File was deleted.',
                    variant: 'success'
                })
            );
        } catch (error) {
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Error deleting file',
                    message: error?.body?.message || error.message,
                    variant: 'error'
                })
            );
            this.showDeleteConfirmModal = false;
        }
    }


    async handleDeleteFile(event) {
        this.isLoading = true;
        const itemId = event.currentTarget.dataset.id;
        const itemName = event.currentTarget.dataset.name;

        console.debug(itemId);

        if (!this.currentDriveId || !itemId) {
            this.error = {
                message: 'Missing required information to delete file.'
            };
            return;
        }
        const driveId = this.currentDriveId;
        console.debug(`item id: ${itemId} driveId: ${driveId} other drive: ${this.currentDriveId}`)

        try {
            await deleteSharepointFile({
                itemId: itemId,
                driveId: driveId
            });
            await this.loadFolderById(this.currentDriveId, this.currentItemId);
            this.showDeleteConfirmModal = false;
            this.deleteFileId = null;
            this.deleteFileName = '';
            this.deleteSuccessMessage(itemName);
        } catch (err) {
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Error deleting file',
                    message: error?.body?.message || error.message,
                    variant: 'error'
                })
            );
            this.showDeleteConfirmModal = false;
        } finally {
            this.isLoading = false;
        }
    }

    /**
     * Drag and drop upload 
     */
    handleDragOver(event) {
        event.preventDefault();
    }

    handleDrop(event) {
        event.preventDefault();
        const file = event.dataTransfer.files[0];
        if (file) {
            // Create a fake event to reuse your existing method
            this.handleUploadFinished({
                target: {
                    files: [file]
                }
            });
        }
    }

    triggerFileInput() {
        this.refs.fileInput.click(); // Uses lwc:ref
    }

    uploadSuccessMessage(fileName) {
        let docName = !!fileName == false ? 'Document' : fileName;
        const evt = new ShowToastEvent({
            title: 'File Upload Succeeded',
            message: `Uploaded ${docName}`,
            variant: 'Success',
            mode: 'dismissible'
        });
        this.dispatchEvent(evt);
    }

    deleteSuccessMessage(fileName) {
        let docName = !!fileName == false ? 'Document' : fileName;
        this.dispatchEvent(
            new ShowToastEvent({
                title: 'File Deleted',
                message: `Deleted ${docName} successfully.`,
                variant: 'info'
            })
        );

    }

    handleCreateFolderOpen() {
        this.newFolderName = '';
        this.showCreateFolderModal = true;
    }

    handleFolderNameChange(event) {
        this.newFolderName = event.target.value;
    }

    handleCreateFolderCancel() {
        this.showCreateFolderModal = false;
    }

    async handleCreateFolder() {
        if (!this.newFolderName || !this.currentDriveId || !this.currentItemId) {
            this.error = {
                message: 'Missing required information to create folder.'
            };
            return;
        }

        try {
            await createFolderInSharePoint({
                folderName: this.newFolderName,
                driveId: this.currentDriveId,
                parentItemId: this.currentItemId
            });

            this.showCreateFolderModal = false;
            await this.loadFolderById(this.currentDriveId, this.currentItemId); // Refresh folder view
        } catch (err) {
            this.error = {
                message: err?.body?.message || err.message
            };
        }
    }


    toggleSearch() {
        this.showSearchInput = !this.showSearchInput;
        if (this.showSearchInput) {
            setTimeout(() => {
                const input = this.template.querySelector('.searchbox-input');
                if (input) input.focus();
            }, 100);
        } else {
            this.searchTerm = '';
            this.refreshData();
        }
    }

    get searchSlideClass() {
        return this.showSearchInput ? 'open' : '';
    }


    handleSearchKey(event) {
        clearTimeout(this.searchTimeout);
        const term = event.target.value.trim();
        console.debug('search term: ', term);
        this.searchTimeout = setTimeout(() => {
            this.searchTerm = term;
            this.runSearch();
        }, 400);
    }

    async runSearch() {
        if (!this.searchTerm) return;
        this.isLoading = true;
        try {
            let results;
            if (this.currentDriveId) {
                results = await searchSharePoint({
                    driveId: this.currentDriveId,
                    searchText: this.searchTerm
                });
            } else {
                // fallback to root folder search
                const raw = await getFolderContentsFromRecord({
                    recordId: this.recordId,
                    objectApiName: this.objectApiName
                });
                results = raw.filter(item => item.name?.toLowerCase().includes(this.searchTerm.toLowerCase()));
            }
            this.files = this.formatFiles(results);
        } catch (err) {
            this.error = {
                message: err?.body?.message || err.message
            };
        } finally {
            this.isLoading = false;
        }
    }


    async handlePreviewClick(event) {
        event.preventDefault();
        this.isLoading = true;
        const anchor = event.currentTarget;
        if (!anchor) {
            console.error('No event.currentTarget');
            return;
        }

        const itemId = anchor.dataset.id;
        const driveId = anchor.dataset.driveid;
        const name = anchor.dataset.name;
        const href = anchor.href;

        console.debug('🔍 handlePreviewClick', JSON.stringify({
            itemId,
            driveId,
            name,
            href
        }, null, '\t'));

        if (!driveId || !itemId) {
            console.warn('❌ Missing driveId or itemId – falling back');
            if (href) window.open(href, '_blank');
            return;
        }

        try {
            const previewUrl = await getSharePointPreviewUrl({
                driveId,
                itemId
            });
            console.debug('Preview URL:', previewUrl);

            if (!previewUrl) {
                throw new Error('Preview URL is undefined');
            }

            this.previewFileName = name;
            this.previewFileUrl = previewUrl;
            this.downloadUrl = href;
            this.showPreviewModal = true;

        } catch (error) {
            console.warn('Error getting preview URL – falling back to tab', error);
            if (href) window.open(href, '_blank');
        } finally {
            this.isLoading = false;
        }
    }

    closePreviewModal() {
        this.showPreviewModal = false;
        this.downloadUrl = '';
        this.previewFileUrl = '';
        this.previewFileName = '';
    }

    formatFiles(data) {
        return data.map(file => {
            if (!file.folder) {
                file.iconName = this.getIconName(file.name);
                file.formattedSize = this.formatSize(file.size);
                file.formattedDate = this.formatDate(file.lastModifiedDateTime);
            }

            // 👇 Add this line for both folder and file cases
            file.driveId = file.parentReference?.driveId;

            return file;
        });
    }


    formatSize(bytes) {
        if (!bytes) return '';
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${sizes[i]}`;
    }

    formatDate(iso) {
        if (!iso) return '';
        const date = new Date(iso);
        return date.toLocaleDateString(undefined, {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    }

    getIconName(name) {
        const ext = (name?.split('.').pop() || '').toLowerCase();
        const map = {
            pdf: 'doctype:pdf',
            doc: 'doctype:word',
            docx: 'doctype:word',
            xls: 'doctype:excel',
            xlsx: 'doctype:excel',
            ppt: 'doctype:ppt',
            pptx: 'doctype:ppt',
            txt: 'doctype:txt',
            jpg: 'doctype:image',
            jpeg: 'doctype:image',
            png: 'doctype:image',
            gif: 'doctype:image',
            zip: 'doctype:zip'
        };
        return map[ext] || 'doctype:unknown';
    }

    /**
     * Upload file formats
     */
    get acceptedFormats() {
        return [
            '.pdf', '.doc', '.docx', '.xls', '.xlsx',
            '.ppt', '.pptx', '.txt', '.jpg', '.jpeg',
            '.png', '.gif', '.zip'
        ];
    }

    get acceptedFormatsJoined() {
        return this.acceptedFormats.join(', ');
    }

    get disableBackButton() {
        // Disable if we’re at root or there's no SharePoint drive loaded
        return this.breadcrumbs.length < 1 || !this.currentDriveId || !this.currentItemId;
    }

    get disableUploadButton() {
        // Disable if there's no resolved SharePoint folder
        return !this.currentDriveId || !this.currentItemId;
    }

    get disableFolderButton() {
        //return !this.currentDriveId;
        return false;
    }

    get disableRefreshButton() {
        //return !this.currentDriveId;
        return false;
    }
}