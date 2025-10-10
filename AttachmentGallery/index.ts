import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { saveAs } from 'file-saver';

interface Attachment {
	documentBody: string;
	mimeType: string;
	title: string;
	noteText: string;
	id: string;
	filename: string;
}

interface IModalState {
	isOpen: boolean,
	isPdfViewerOpen: boolean
}

interface IPdfState {
	pdf: any,
	currentPage: number,
	zoom: number
}

export class AttachmentGallery implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _previewImg: HTMLImageElement;
	private _thumbnailsGallery: HTMLDivElement;
	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _currentIndex: number;
	private _notes: Attachment[];
	private _modalContainer: HTMLDivElement;
	private _modalContentContainer: HTMLDivElement;
	private _modalHeaderText: HTMLHeadingElement;
	private _modalImage: HTMLImageElement;
	private _noteText: HTMLParagraphElement;
	private _noteTextContainer: HTMLDivElement;
	private _pdfCanvas: HTMLCanvasElement;
	private pdfState: IPdfState;
	private modalState: IModalState;
	private _pdfViewerContainer: HTMLDivElement;
	private _modalImageContainer: HTMLDivElement;
	private _pdfPageInput: HTMLInputElement;
	private _pdfTotalPages: HTMLSpanElement;
	private _pdfPageControlsContainer: HTMLDivElement;
	private _imageViewerContainer: HTMLDivElement;
	private pdfImageSrc: string;
	private _mainDivContainer: HTMLDivElement;
	private _notFoundContainer: HTMLDivElement;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._notes = [];
		this._context = context;
		this.setPreview = this.setPreview.bind(this);
		this.openModal = this.openModal.bind(this);
		this.closeModal = this.closeModal.bind(this);
		this.changeImage = this.changeImage.bind(this);
		this.setPreviewFromThumbnail = this.setPreviewFromThumbnail.bind(this);
		this.toggleNoteColumn = this.toggleNoteColumn.bind(this);
		this.generateImageSrcUrl = this.generateImageSrcUrl.bind(this);
		//this.GetAttachmentsDemo = this.GetAttachmentsDemo.bind(this);
		this.b64toBlob = this.b64toBlob.bind(this);
		this.downloadFile = this.downloadFile.bind(this);
		this.setPdfViewer = this.setPdfViewer.bind(this);
		this.pdfRender = this.pdfRender.bind(this);
		this.togglePdfViwer = this.togglePdfViwer.bind(this);
		this.setModalImage = this.setModalImage.bind(this);
		this.changePdfPage = this.changePdfPage.bind(this);
		this.zoomPdfPage = this.zoomPdfPage.bind(this);
		this.setPdfPage = this.setPdfPage.bind(this);

		this._context.resources.getResource('img/pdf_icon.png',
						this.setPdfImage.bind(this), this.showError.bind(this,"ERROR with PDF Image!"));

		this.modalState = {
			isOpen: false,
			isPdfViewerOpen: false
		}

		this.pdfState = {
			pdf: null,
			currentPage: 1,
			zoom: 1
		}

		this._container = document.createElement('div');

		//--------- Attachments not found placeholder
		this._notFoundContainer = document.createElement('div');

		let refreshIcon = document.createElement('i');
		refreshIcon.className = 'dwc-top-right ms-Icon ms-Icon--Refresh';

		this._notFoundContainer.appendChild(refreshIcon);

		let notFoundText = document.createElement('p');
		notFoundText.classList.add('dwc-center');
		notFoundText.innerText = 'Attachments not found. Press Refresh to try to load them again.';

		this._notFoundContainer.appendChild(notFoundText);

		this._container.appendChild(this._notFoundContainer);

		let mainContainer = document.createElement('div');
		mainContainer.classList.add('main-container','dwc-hide');

		this._mainDivContainer = mainContainer;

		//-------- creating thumbnails
		this._thumbnailsGallery = document.createElement("div");
		this._thumbnailsGallery.classList.add('thumbnailsList');

		//-------- creating preview section
		let bigPreview = document.createElement("div");
		bigPreview.classList.add('preview-section');

		this._previewImg = document.createElement('img');
		this._previewImg.classList.add('preview-img');
		this._previewImg.onclick = () => this.openModal();
		bigPreview.appendChild(this._previewImg);

		//-------- prev and next buttons
		let next = document.createElement('a');
		next.classList.add('arrow-button', 'preview-next');
		next.innerHTML = "&#10095;";
		next.onclick = () => this.changeImage(1);

		let prev = document.createElement('a');
		prev.classList.add('arrow-button', 'preview-prev');
		prev.innerHTML = "&#10094;";
		prev.onclick = () => this.changeImage(-1);

		bigPreview.appendChild(prev);
		bigPreview.appendChild(next);

		mainContainer.appendChild(bigPreview);
		mainContainer.appendChild(this._thumbnailsGallery);

		this._container.appendChild(mainContainer);

		container.appendChild(this._container);

		//--------- create modal

		this._modalContainer = document.createElement('div');
		this._modalContainer.classList.add('dwc-modal');

		let modalContent = document.createElement('div');
		modalContent.classList.add('dwc-modal-content');

		this._modalContentContainer = modalContent;

		//--------- create modal header
		let modalHeader = document.createElement('div');
		modalHeader.classList.add('dwc-modal-header');

		//--------- create pdf controls
		let pdfControlsContainer = document.createElement('div');

		pdfControlsContainer.classList.add('dwc-flex-container', 'dwc-pageinput-container', 'dwc-hide');

		let prevPdfPageIcon = document.createElement('i');
		prevPdfPageIcon.className = "header-icon ms-Icon ms-Icon--ChevronLeft";
		prevPdfPageIcon.addEventListener('click', () => this.changePdfPage(-1));

		pdfControlsContainer.appendChild(prevPdfPageIcon);

		let pdfPageNumberContainer = document.createElement('div');

		let pdfPageInput = document.createElement('input');
		pdfPageInput.value = '1';
		pdfPageInput.className = 'dwc-page-input';
		pdfPageInput.addEventListener('keypress', (e) => this.setPdfPage(e));

		pdfPageNumberContainer.appendChild(pdfPageInput);

		this._pdfPageInput = pdfPageInput;

		let totalPagesSpan = document.createElement('span');
		totalPagesSpan.innerHTML = " / 0";
		totalPagesSpan.className = 'dwc-page-span';

		pdfPageNumberContainer.appendChild(totalPagesSpan);

		this._pdfTotalPages = totalPagesSpan;

		pdfControlsContainer.appendChild(pdfPageNumberContainer);

		let nextPdfPageIcon = document.createElement('i');
		nextPdfPageIcon.className = "header-icon ms-Icon ms-Icon--ChevronRight";
		nextPdfPageIcon.addEventListener('click', () => this.changePdfPage(1));

		pdfControlsContainer.appendChild(nextPdfPageIcon);

		let zoomControlsContainer = document.createElement('div');
		zoomControlsContainer.className = 'dwc-zoom-container';

		let zoomText = document.createElement('span');
		zoomText.style.color = 'white';
		zoomText.innerHTML = 'Zoom: ';

		let plusIcon = document.createElement('i');
		plusIcon.className = "header-icon dwc-side-margin-4 ms-Icon ms-Icon--Add";
		plusIcon.addEventListener('click', () => this.zoomPdfPage(0.2));

		let minusIcon = document.createElement('i');
		minusIcon.className = "header-icon dwc-side-margin-4 ms-Icon ms-Icon--Remove";
		minusIcon.addEventListener('click', () => this.zoomPdfPage(-0.2));

		zoomControlsContainer.appendChild(zoomText);
		zoomControlsContainer.appendChild(plusIcon);
		zoomControlsContainer.appendChild(minusIcon);

		pdfControlsContainer.appendChild(zoomControlsContainer);

		this._pdfPageControlsContainer = pdfControlsContainer;

		//---------- create modal buttons
		let rightHeaderContainer = document.createElement('div');

		rightHeaderContainer.classList.add("header-icon-container");

		let downloadIcon = document.createElement('i');
		downloadIcon.className = "header-icon ms-Icon ms-Icon--Download";
		downloadIcon.addEventListener('click', this.downloadFile);

		rightHeaderContainer.appendChild(downloadIcon);

		let infoIcon = document.createElement('i');
		infoIcon.className = "header-icon ms-Icon ms-Icon--Info";
		infoIcon.addEventListener('click', this.toggleNoteColumn);

		rightHeaderContainer.appendChild(infoIcon);

		let closeIcon = document.createElement('i');
		closeIcon.className = "header-icon ms-Icon ms-Icon--ChromeClose";
		closeIcon.addEventListener('click', this.closeModal);

		rightHeaderContainer.appendChild(closeIcon);

		//--------- create modal header text
		let leftHeaderContainer = document.createElement('div');

		let headerText = document.createElement('h3');
		headerText.innerHTML = "0/10";
		this._modalHeaderText = headerText;

		leftHeaderContainer.appendChild(headerText);

		modalHeader.appendChild(leftHeaderContainer);
		modalHeader.appendChild(pdfControlsContainer);
		modalHeader.appendChild(rightHeaderContainer);

		modalContent.appendChild(modalHeader);

		//--------- create modal body

		let modalBody = document.createElement('div');
		modalBody.classList.add('dwc-modal-body');

		//--------- add prev/next buttons

		let nextImgModal = document.createElement('a');
		nextImgModal.classList.add('arrow-button', 'preview-next');
		nextImgModal.innerHTML = "&#10095;";
		nextImgModal.onclick = () => this.changeImage(1);

		let prevImgModal = document.createElement('a');
		prevImgModal.classList.add('arrow-button', 'preview-prev');
		prevImgModal.innerHTML = "&#10094;";
		prevImgModal.onclick = () => this.changeImage(-1);

		modalBody.appendChild(nextImgModal);
		modalBody.appendChild(prevImgModal);

		//--------- image container

		let imageViewerContainer = document.createElement('div');
		imageViewerContainer.className = 'dwc-flex-container';

		this._imageViewerContainer = imageViewerContainer;

		modalBody.appendChild(imageViewerContainer);

		let modalImageContainer = document.createElement('div');
		modalImageContainer.classList.add('dwc-modal-img-container');

		this._modalImage = document.createElement('img');
		this._modalImage.classList.add('dwc-modal-img');

		modalImageContainer.appendChild(this._modalImage);
		imageViewerContainer.appendChild(modalImageContainer);

		this._modalImageContainer = modalImageContainer;

		//-------- create preview note container

		let modalNoteTextContainer = document.createElement('div');
		modalNoteTextContainer.classList.add('dwc-modal-notetext-container', 'dwc-hide');

		let noteText = document.createElement('p');
		noteText.className = 'dwc-modal-notetext';
		modalNoteTextContainer.appendChild(noteText);

		this._noteText = noteText;
		this._noteTextContainer = modalNoteTextContainer;

		imageViewerContainer.appendChild(modalNoteTextContainer);

		//--------- create pdf viewer container

		let pdfViewerContainer = document.createElement('div');
		pdfViewerContainer.classList.add('dwc-hide');

		let canvasContainer = document.createElement('div');
		canvasContainer.className = 'dwc-pdf-canvas-container';
		canvasContainer.id = "canvas_container";

		this._pdfCanvas = document.createElement('canvas');
		this._pdfCanvas.id = "pdf_renderer";

		canvasContainer.appendChild(this._pdfCanvas);

		pdfViewerContainer.appendChild(canvasContainer);

		modalBody.appendChild(pdfViewerContainer);

		this._pdfViewerContainer = pdfViewerContainer;

		modalContent.appendChild(modalBody);
		this._modalContainer.appendChild(modalContent);

		document.body.appendChild(this._modalContainer);

		console.log("context", context);

		let curentRecord: ComponentFramework.EntityReference = {
			id: (<any>context).page.entityId,
			name: (<any>context).page.entityTypeName
		}

		console.log("curentRecord", curentRecord);

		const pdfScript = document.createElement('script');
		pdfScript.src = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js';

		document.body.appendChild(pdfScript);

		this.GetAttachments(curentRecord).then(result => this.CreateGallery(result));
    }

    /**
    * Generate Image Element src url
    * @param fileType file extension
    * @param fileContent file content, base 64 format
    */
    private generateImageSrcUrl(fileType: string, fileContent: string): string {
        return "data:" + fileType + ";base64, " + fileContent;
    }

    private toggleNoteColumn(): void {
        if (this.modalState.isPdfViewerOpen) return;

        if (this._noteTextContainer.classList.contains('dwc-hide')) {
            this._noteTextContainer.classList.remove('dwc-hide');
        } else {
            this._noteTextContainer.classList.add('dwc-hide');
        }
    }

    private openModal(): void {
        this._modalContainer.style.display = "block";
        this.modalState.isOpen = true;
        let currentNote = this._notes[this._currentIndex];

        this.setModalImage(currentNote);
    }

    private closeModal(): void {
        this._modalContainer.style.display = "none";
        this.modalState.isOpen = false;
    }

    private async GetAttachments(curentRecord: ComponentFramework.EntityReference): Promise<Attachment[]> {
        console.log("GetAttachments started");

        const searchQuery = "?$select=annotationid,documentbody,mimetype,notetext,subject,filename" +
            "&$filter=_objectid_value eq " +
            curentRecord.id +
            " and isdocument eq true and (startswith(mimetype, 'application/pdf') or startswith(mimetype, 'image/'))";
        try {

            const result = await this._context.webAPI.retrieveMultipleRecords("annotation", searchQuery);

            console.log("Retrieved attachments", result);
            if (result && result.entities) {
                for (let index = 0; index < result.entities.length; index++) {

                    let item: Attachment = {
                        id: result.entities[index].id,
                        mimeType: result.entities[index].mimetype,
                        noteText: result.entities[index].notetext,
                        title: result.entities[index].subject || result.entities[index].filename,
                        filename: result.entities[index].filename,
                        documentBody: result.entities[index].documentbody
                    };
                    this._notes.push(item);
                }
            }
        } catch (error) {
            console.error("ERROR RETRIEVING ATTACHMENT");
            console.error(error);
        }

        return this._notes;
    }

    private setPdfImage(body:string){
        this.pdfImageSrc = this.generateImageSrcUrl('image/png',body);
    }

    private showError(text:string){
        console.error("ERROR:", text);
    }

    private CreateGallery(result: Attachment[]): any {
        console.log("Create Gallery Started");
        console.log("Notes: ", result);
        if (result.length > 0) {
            let count = 0;
            for (let i = 0; i < result.length; i++) {
                let newImg = document.createElement('img');
                newImg.className = "thumbnail";
                newImg.src = result[i].mimeType.indexOf('pdf') == -1
                    ? this.generateImageSrcUrl(result[i].mimeType, result[i].documentBody)
                    : this.pdfImageSrc;
                newImg.alt = count.toString();
                newImg.addEventListener('click', this.setPreviewFromThumbnail);
                this._thumbnailsGallery.appendChild(newImg);
                count++;
            }

            this._notFoundContainer.classList.add('dwc-hide');
            this._mainDivContainer.classList.remove('dwc-hide');

            this._currentIndex = 0;
            this.setPreview(0);
        } 
    }

    private setPreviewFromThumbnail(e: Event) {
        let currentImgIndex = parseInt((e.srcElement as HTMLImageElement).alt);
        this.setPreview(currentImgIndex);
    }

    private setModalImage(note: Attachment) {
        let isAttachmentPdf = note.mimeType.indexOf('pdf') != -1;

        if (this.modalState.isOpen && !this.modalState.isPdfViewerOpen && isAttachmentPdf) {
            this.togglePdfViwer(true);
            this.setPdfViewer(note);
        } else {
            if (this.modalState.isOpen && this.modalState.isPdfViewerOpen && isAttachmentPdf) {
                this.setPdfViewer(note);
            } else {
                if (this.modalState.isOpen && this.modalState.isPdfViewerOpen && !isAttachmentPdf) {
                    this.togglePdfViwer(false);
                }
                this._modalImage.src = this._previewImg.src;
            }
        }
    }

    private setPreview(currentNoteNumber: number) {
        if (currentNoteNumber === this._notes.length) { currentNoteNumber = 0 }
        if (currentNoteNumber < 0) { currentNoteNumber = this._notes.length - 1 }
        let currentNote = this._notes[currentNoteNumber];

        let isAttachmentPdf = currentNote.mimeType.indexOf('pdf') != -1;

        this._previewImg.src = !isAttachmentPdf
            ? this.generateImageSrcUrl(currentNote.mimeType, currentNote.documentBody)
            : this.pdfImageSrc;

        this._modalHeaderText.innerHTML = (currentNoteNumber + 1).toString() + " / " + this._notes.length.toString()
            + " " + currentNote.title;

        if (this.modalState.isOpen) {
            this.setModalImage(currentNote);
        }

        this._noteText.innerHTML = currentNote.noteText;
        this._currentIndex = currentNoteNumber;
    }

    private changeImage(moveIndex: number) {
        this.setPreview(this._currentIndex += moveIndex);
    }

    //=========== FILE DOWNLOAD LOGIC ===============

    private b64toBlob(b64Data: string, contentType: string, sliceSize: number): Blob {
        contentType = contentType || '';
        sliceSize = sliceSize || 512;

        var byteCharacters = atob(b64Data);
        var byteArrays = [];

        for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
            var slice = byteCharacters.slice(offset, offset + sliceSize);

            var byteNumbers = new Array(slice.length);
            for (var i = 0; i < slice.length; i++) {
                byteNumbers[i] = slice.charCodeAt(i);
            }

            var byteArray = new Uint8Array(byteNumbers);

            byteArrays.push(byteArray);
        }

        var blob = new Blob(byteArrays, { type: contentType });
        return blob;
    }

    private downloadFile(): void {
        let note = this._notes[this._currentIndex];
        let blob = this.b64toBlob(note.documentBody, note.mimeType, 512);
        saveAs(blob, note.filename);
    }


    //=========== PDF LOGIC

    private setPdfViewer(pdfItem: Attachment) {
        let pdfData = atob(pdfItem.documentBody);
        // @ts-ignore
        pdfjsLib.getDocument({ data: pdfData }).promise.then((pdf) => {
            this.pdfState.pdf = pdf;
            this.pdfState.currentPage = 1;
            this.pdfState.zoom = 1;
            this._pdfPageInput.value = '1';

            this._pdfTotalPages.innerHTML = " / " + pdf._pdfInfo.numPages.toString();
            this.pdfRender();
        });
    }

    private pdfRender() {
        // @ts-ignore
        this.pdfState.pdf.getPage(this.pdfState.currentPage).then((page) => {

            let canvas = this._pdfCanvas;
            var ctx = canvas.getContext('2d');
            var viewport = page.getViewport({
                scale: this.pdfState.zoom
            });

            canvas.width = viewport.width;
            canvas.height = viewport.height

            page.render({
                canvasContext: ctx,
                viewport: viewport
            });
        });
    }

    private togglePdfViwer(visible: boolean): void {
        if (visible) {
            this._pdfViewerContainer.classList.remove("dwc-hide");
            this._pdfPageControlsContainer.classList.remove("dwc-hide");
            this._imageViewerContainer.classList.add("dwc-hide");
            this._modalContentContainer.classList.add('dwc-overflow');
            this.modalState.isPdfViewerOpen = true;
        } else {
            this._pdfViewerContainer.classList.add("dwc-hide");
            this._pdfPageControlsContainer.classList.add("dwc-hide");
            this._imageViewerContainer.classList.remove("dwc-hide");
            this._modalContentContainer.classList.remove('dwc-overflow');
            this.modalState.isPdfViewerOpen = false;
        }
    }

    private setPdfPage(event: KeyboardEvent): void {
        if (this.pdfState.pdf == null) return;

        // Get key code
        let code = (event.keyCode ? event.keyCode : event.which);

        // If key code matches that of the Enter key
        if (code == 13) {
            let desiredPage = parseInt(this._pdfPageInput.value);

            if (desiredPage >= 1 &&
                desiredPage <= this.pdfState.pdf._pdfInfo.numPages) {

                this.pdfState.currentPage = desiredPage;
                this.pdfRender();
            }
        }
    }

    private changePdfPage(pageShift: number): void {
        if (this.pdfState.pdf == null) return;

        let nextPage = this.pdfState.currentPage;
        nextPage += pageShift;

        if (nextPage <= 0 || nextPage > this.pdfState.pdf._pdfInfo.numPages) return;

        this.pdfState.currentPage = nextPage;
        this._pdfPageInput.value = this.pdfState.currentPage.toString();
        this.pdfRender();
    }

    private zoomPdfPage(zoomChange: number): void {
        if (this.pdfState.pdf == null) return;
        this.pdfState.zoom += zoomChange;
        this.pdfRender();
    }

    //---------- END PDF LOGIC

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}
