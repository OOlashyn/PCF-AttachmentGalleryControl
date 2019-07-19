import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { img1, img2 } from "./demoImages";
import { saveAs } from 'file-saver';
var pdfjsLib = require('pdfjs-dist');

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

export class AttachmentGalleryControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _previewImg: HTMLImageElement;
	private _thumbnailsGallery: HTMLDivElement;
	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _currentIndex: number;
	private _notes: Attachment[];
	private _modalContainer: HTMLDivElement;
	private _modalHeaderText: HTMLHeadingElement;
	private _modalImage: HTMLImageElement;
	private _noteTitle: HTMLParagraphElement;
	private _noteText: HTMLParagraphElement;
	private _textColumn: HTMLDivElement;
	private _imageColumn: HTMLDivElement;
	private _infoImg: HTMLImageElement;
	private _pdfCanvas: HTMLCanvasElement;
	private pdfState: IPdfState;
	private modalState: IModalState;
	private _pdfViewerContainer: HTMLDivElement;
	private _modalImageConatiner: HTMLDivElement;
	private _pdfPageInput: HTMLInputElement;
	private _pdfTotalPages: HTMLSpanElement;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._notes = [];
		this._context = context;
		this.setPreview = this.setPreview.bind(this);
		this.openModal = this.openModal.bind(this);
		this.closeModal = this.closeModal.bind(this);
		this.changeImage = this.changeImage.bind(this);
		this.setPreviewFromThumbnail = this.setPreviewFromThumbnail.bind(this);
		this.toggleColumn = this.toggleColumn.bind(this);
		this.generateImageSrcUrl = this.generateImageSrcUrl.bind(this);
		this.GetAttachmentsDemo = this.GetAttachmentsDemo.bind(this);
		this.b64toBlob = this.b64toBlob.bind(this);
		this.downloadFile = this.downloadFile.bind(this);
		this.setPdfViewer = this.setPdfViewer.bind(this);
		this.pdfRender = this.pdfRender.bind(this);
		this.togglePdfViwer = this.togglePdfViwer.bind(this);
		this.setModalImage = this.setModalImage.bind(this);
		this.changePdfPage = this.changePdfPage.bind(this);

		this.modalState = {
			isOpen: false,
			isPdfViewerOpen: false
		}

		this.pdfState = {
			pdf: null,
			currentPage: 1,
			zoom: 1
		}

		this._container = document.createElement("div");

		let mainContainer = document.createElement('div');
		mainContainer.classList.add('main-container');

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

		//--------- create modal header
		let modalHeader = document.createElement('div');
		modalHeader.classList.add('dwc-modal-header');

		//--------- create pdf controls
		let pdfControlsContainer = document.createElement('div');

		pdfControlsContainer.classList.add('dwc-flex-container','dwc-pageinput-container');

		let prevPdfPageIcon = document.createElement('i');
		prevPdfPageIcon.className = "header-icon ms-Icon ms-Icon--ChevronLeft";
		prevPdfPageIcon.addEventListener('click', () => this.changePdfPage(-1));

		pdfControlsContainer.appendChild(prevPdfPageIcon);

		let pdfPageNumberContainer = document.createElement('div');
		
		let pdfPageInput = document.createElement('input');
		pdfPageInput.value = '1';
		pdfPageInput.className = 'dwc-page-input';

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

		//---------- create modal buttons
		let rightHeaderContainer = document.createElement('div');

		rightHeaderContainer.classList.add("header-icon-container");

		let downloadIcon = document.createElement('i');
		downloadIcon.className = "header-icon ms-Icon ms-Icon--Download";
		downloadIcon.addEventListener('click', this.downloadFile);

		rightHeaderContainer.appendChild(downloadIcon);

		let infoIcon = document.createElement('i');
		infoIcon.className = "header-icon ms-Icon ms-Icon--Info";
		infoIcon.addEventListener('click', this.toggleColumn);

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

		let modalImageConatiner = document.createElement('div');
		modalImageConatiner.classList.add('dwc-modal-img-container');

		this._modalImage = document.createElement('img');
		this._modalImage.classList.add('dwc-modal-img');

		modalImageConatiner.appendChild(this._modalImage);
		modalBody.appendChild(modalImageConatiner);

		this._modalImageConatiner = modalImageConatiner;

		//-------- create preview note container

		let modalNoteTextContainer = document.createElement('div');
		modalNoteTextContainer.classList.add('dwc-modal-notetext', 'dwc-hidden');

		let noteText = document.createElement('p');
		modalNoteTextContainer.appendChild(noteText);

		this._noteText = noteText;

		modalBody.appendChild(modalNoteTextContainer);

		//--------- create pdf viewer container

		let pdfViewerContainer = document.createElement('div');
		pdfViewerContainer.classList.add('dwc-hide');

		let navigationPanel = document.createElement('div');

		let canvasContainer = document.createElement('div');
		canvasContainer.id = "canvas_container";

		this._pdfCanvas = document.createElement('canvas');
		this._pdfCanvas.id = "pdf_renderer";

		canvasContainer.appendChild(this._pdfCanvas);

		pdfViewerContainer.appendChild(navigationPanel);
		pdfViewerContainer.appendChild(canvasContainer);

		modalBody.appendChild(pdfViewerContainer);

		this._pdfViewerContainer = pdfViewerContainer;


		modalContent.appendChild(modalBody);
		this._modalContainer.appendChild(modalContent);

		document.body.appendChild(this._modalContainer);

		let curentRecord: ComponentFramework.EntityReference = {
			id: (<any>context).page.entityId,
			name: (<any>context).page.entityTypeName
		}

		//this.GetAttachments(curentRecord).then( result => this.CreateGallery(result));
		this.GetAttachmentsDemo();
	}



	private GetAttachmentsDemo(): void {

		for (let index = 0; index < 10; index++) {

			let item: Attachment = {
				id: index.toString(),
				mimeType: "jpeg",
				noteText: "Text for image number: " + index.toString(),
				title: "Title " + index.toString(),
				filename: "img" + index.toString() + ".jpeg",
				documentBody: index % 2 ? img1 : img2
			};
			this._notes.push(item);
		}

		let pdfItem: Attachment = {
			id: '10',
			mimeType: "pdf",
			noteText: "Text for image number: " + '10',
			title: "Title " + '10',
			filename: "pdf_10.pdf",
			documentBody: 'JVBERi0xLjcKCjEgMCBvYmogICUgZW50cnkgcG9pbnQKPDwKICAvVHlwZSAvQ2F0YWxvZwog' +
				'IC9QYWdlcyAyIDAgUgo+PgplbmRvYmoKCjIgMCBvYmoKPDwKICAvVHlwZSAvUGFnZXMKICAv' +
				'TWVkaWFCb3ggWyAwIDAgMjAwIDIwMCBdCiAgL0NvdW50IDEKICAvS2lkcyBbIDMgMCBSIF0K' +
				'Pj4KZW5kb2JqCgozIDAgb2JqCjw8CiAgL1R5cGUgL1BhZ2UKICAvUGFyZW50IDIgMCBSCiAg' +
				'L1Jlc291cmNlcyA8PAogICAgL0ZvbnQgPDwKICAgICAgL0YxIDQgMCBSIAogICAgPj4KICA+' +
				'PgogIC9Db250ZW50cyA1IDAgUgo+PgplbmRvYmoKCjQgMCBvYmoKPDwKICAvVHlwZSAvRm9u' +
				'dAogIC9TdWJ0eXBlIC9UeXBlMQogIC9CYXNlRm9udCAvVGltZXMtUm9tYW4KPj4KZW5kb2Jq' +
				'Cgo1IDAgb2JqICAlIHBhZ2UgY29udGVudAo8PAogIC9MZW5ndGggNDQKPj4Kc3RyZWFtCkJU' +
				'CjcwIDUwIFRECi9GMSAxMiBUZgooSGVsbG8sIHdvcmxkISkgVGoKRVQKZW5kc3RyZWFtCmVu' +
				'ZG9iagoKeHJlZgowIDYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDEwIDAwMDAwIG4g' +
				'CjAwMDAwMDAwNzkgMDAwMDAgbiAKMDAwMDAwMDE3MyAwMDAwMCBuIAowMDAwMDAwMzAxIDAw' +
				'MDAwIG4gCjAwMDAwMDAzODAgMDAwMDAgbiAKdHJhaWxlcgo8PAogIC9TaXplIDYKICAvUm9v' +
				'dCAxIDAgUgo+PgpzdGFydHhyZWYKNDkyCiUlRU9G'
		};
		this._notes.push(pdfItem);

		//this.setPdfViewer(pdfItem);

		this.CreateGallery(this._notes);
	}

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

	private setImage(fileType: string, fileContent: string): void {
		console.log(fileContent);
		let imageUrl: string = this.generateImageSrcUrl(fileType, fileContent);
		this._infoImg.src = imageUrl;
	}

	/**
	 * Generate Image Element src url
	 * @param fileType file extension
	 * @param fileContent file content, base 64 format
	 */
	private generateImageSrcUrl(fileType: string, fileContent: string): string {
		return "data:image/" + fileType + ";base64, " + fileContent;
	}

	private toggleColumn(): void {
		//
		//this._noteText.parentElement.classList.remove('dwc-hide');
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
		let searchQuery = "?$select=annotationid,documentbody,mimetype,notetext,subject,filename&$filter=_objectid_value eq " +
			curentRecord.id +
			" and  isdocument eq true and startswith(mimetype, 'image/')";
		const result = await this._context.webAPI.retrieveMultipleRecords("annotation", searchQuery);

		if (result && result.entities) {
			for (let index = 0; index < result.entities.length; index++) {

				let item: Attachment = {
					id: result.entities[index].id,
					mimeType: result.entities[index].mimetype,
					noteText: result.entities[index].notetext,
					title: result.entities[index].subject,
					filename: result.entities[index].filename,
					documentBody: result.entities[index].documentbody
				};
				this._notes.push(item);
			}
		}

		return this._notes;
	}

	private CreateGallery(result: Attachment[]): any {
		if (result.length > 0) {
			let count = 0;
			for (let i = 0; i < result.length; i++) {
				let newImg = document.createElement('img');
				newImg.className = "thumbnail";
				newImg.src = result[i].mimeType != "pdf"
					? this.generateImageSrcUrl(result[i].mimeType, result[i].documentBody)
					: 'img/pdf_icon.png';
				newImg.alt = count.toString();
				newImg.addEventListener('click', this.setPreviewFromThumbnail);
				this._thumbnailsGallery.appendChild(newImg);
				count++;
			}

			this._currentIndex = 0;
			this.setPreview(0);
		}
	}

	private setPreviewFromThumbnail(e: Event) {
		let currentImgIndex = parseInt((e.srcElement as HTMLImageElement).alt);
		this.setPreview(currentImgIndex);
	}

	private setModalImage(note:Attachment){
		let isAttachmentPdf = note.mimeType == "pdf";

		if (this.modalState.isOpen && !this.modalState.isPdfViewerOpen && isAttachmentPdf) {
			this.togglePdfViwer(true);
			this.setPdfViewer(note);
		} else {
			if (this.modalState.isOpen && this.modalState.isPdfViewerOpen && !isAttachmentPdf) {
				this.togglePdfViwer(false);
			}
			this._modalImage.src = this._previewImg.src;
		}
	}

	private setPreview(currentNoteNumber: number) {
		if (currentNoteNumber === this._notes.length) { currentNoteNumber = 0 }
		if (currentNoteNumber < 0) { currentNoteNumber = this._notes.length - 1 }
		let currentNote = this._notes[currentNoteNumber];

		let isAttachmentPdf = currentNote.mimeType == "pdf";

		this._previewImg.src = !isAttachmentPdf
			? this.generateImageSrcUrl(currentNote.mimeType, currentNote.documentBody)
			: 'img/pdf_icon.png';

		this._modalHeaderText.innerHTML = (currentNoteNumber + 1).toString() + " / " + this._notes.length.toString()
			+ " " + currentNote.title;

		if(this.modalState.isOpen){
			this.setModalImage(currentNote);
		}

		this._noteText.innerHTML = currentNote.noteText;
		this._currentIndex = currentNoteNumber;
	}

	private changeImage(moveIndex: number) {
		this.setPreview(this._currentIndex += moveIndex);
	}

	//---------- PDF LOGIC

	private setPdfViewer(pdfItem: Attachment) {
		let pdfData = atob(pdfItem.documentBody);
		// @ts-ignore
		pdfjsLib.getDocument({ data: pdfData }).promise.then((pdf) => {
			this.pdfState.pdf = pdf;
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
			this._modalImageConatiner.classList.add("dwc-hide");
			this.modalState.isPdfViewerOpen = true;
		} else {
			this._pdfViewerContainer.classList.add("dwc-hide");
			this._modalImageConatiner.classList.remove("dwc-hide");
			this.modalState.isPdfViewerOpen = false;
		}
	}

	private changePdfPage(pageShift:number): void{
		if (this.pdfState.pdf == null || this.pdfState.currentPage == 1) return;
		this.pdfState.currentPage += pageShift;
		this._pdfPageInput.value = this.pdfState.currentPage.toString();
		this.pdfRender();
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		document.body.removeChild(this._modalContainer);
	}
}