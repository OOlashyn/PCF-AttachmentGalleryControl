import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { img1, img2 } from "./demoImages";

interface Attachment {
	documentBody: string;
	mimeType: string;
	title: string;
	noteText: string;
	id: string;
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

		let leftHeaderContainer = document.createElement('div');
		let rightHeaderContainer = document.createElement('div');

		let closeSpan = document.createElement('span');
		closeSpan.innerHTML = "&times;";
		closeSpan.classList.add('close-span');
		closeSpan.onclick = () => this.closeModal();
		
		let headerText = document.createElement('h3');
		headerText.innerHTML = "0/10";
		this._modalHeaderText = headerText;

		rightHeaderContainer.appendChild(closeSpan);
		leftHeaderContainer.appendChild(headerText);

		modalHeader.appendChild(leftHeaderContainer);
		modalHeader.appendChild(rightHeaderContainer);

		modalContent.appendChild(modalHeader);
		
		//--------- create modal body

		let modalBody = document.createElement('div');
		modalBody.classList.add('dwc-modal-body');
		
		let modalImageConatiner = document.createElement('div');
		modalImageConatiner.classList.add('dwc-modal-img-container');

		let nextImgModal = document.createElement('a');
		nextImgModal.classList.add('arrow-button', 'preview-next');
		nextImgModal.innerHTML = "&#10095;";
		nextImgModal.onclick = () => this.changeImage(1);

		let prevImgModal = document.createElement('a');
		prevImgModal.classList.add('arrow-button', 'preview-prev');
		prevImgModal.innerHTML = "&#10094;";
		prevImgModal.onclick = () => this.changeImage(-1);

		modalImageConatiner.appendChild(nextImgModal);
		modalImageConatiner.appendChild(prevImgModal);

		this._modalImage = document.createElement('img');
		this._modalImage.classList.add('dwc-modal-img');
		
		modalImageConatiner.appendChild(this._modalImage);
		modalBody.appendChild(modalImageConatiner);

		modalContent.appendChild(modalBody);
		this._modalContainer.appendChild(modalContent);

		document.body.appendChild(this._modalContainer);

		let curentRecord: ComponentFramework.EntityReference = {
			id: (<any>context).page.entityId,
			name: (<any>context).page.entityTypeName
		}

		this.GetAttachments(curentRecord).then( result => this.CreateGallery(result));
		//this.GetAttachmentsDemo();
	}

	private GetAttachmentsDemo(): void {

		for (let index = 0; index < 10; index++) {

			let item: Attachment = {
				id: index.toString(),
				mimeType: "jpeg",
				noteText: "Text for image number: " + index.toString(),
				title: "Title " + index.toString(),
				documentBody: index % 2 ? img1 : img2
			};
			this._notes.push(item);
		}

		this.CreateGallery(this._notes);
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
		this._textColumn.classList.toggle('w3-hide');
		this._textColumn.classList.toggle('w3-rest');
		this._imageColumn.classList.toggle('w3-threequarter');
	}

	private openModal(): void {
		this._modalContainer.style.display = "block";
	}

	private closeModal(): void {
		this._modalContainer.style.display = "none";
	}

	private async GetAttachments(curentRecord: ComponentFramework.EntityReference): Promise<Attachment[]> {
		let searchQuery = "?$select=annotationid,documentbody,mimetype,notetext,subject&$filter=_objectid_value eq " +
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
				newImg.src = this.generateImageSrcUrl(result[i].mimeType, result[i].documentBody);
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

	private setPreview(currentNoteNumber: number) {
		if (currentNoteNumber === this._notes.length) { currentNoteNumber = 0 }
		if (currentNoteNumber < 0) { currentNoteNumber = this._notes.length - 1 }
		this._previewImg.src = this.generateImageSrcUrl(this._notes[currentNoteNumber].mimeType,
			this._notes[currentNoteNumber].documentBody);
		this._modalHeaderText.innerHTML = (currentNoteNumber+1).toString() + " / " + this._notes.length.toString();
		this._modalImage.src = this._previewImg.src;
		// this._noteTitle.innerHTML = this._notes[currentNoteNumber].title;
		// this._noteText.innerHTML = this._notes[currentNoteNumber].noteText;
		this._currentIndex = currentNoteNumber;
	}

	private changeImage(moveIndex: number) {
		this.setPreview(this._currentIndex += moveIndex);
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
	}
}