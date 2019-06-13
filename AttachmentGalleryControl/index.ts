import {IInputs, IOutputs} from "./generated/ManifestTypes";

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
	private _modalImage: HTMLImageElement;
	private _noteTitle: HTMLParagraphElement;
	private _noteText: HTMLParagraphElement;
	private _textColumn: HTMLDivElement;
	private _imageColumn: HTMLDivElement;

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
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this._notes = [];
		this._context = context;
		this.setPreview = this.setPreview.bind(this);
		this.openModal = this.openModal.bind(this);
		this.closeModal = this.closeModal.bind(this);
		this.changeImage = this.changeImage.bind(this);
		this.setPreviewFromThumbnail = this.setPreviewFromThumbnail.bind(this);
		this.toggleColumn = this.toggleColumn.bind(this);

		this._container = document.createElement("div");

		let mainSection = document.createElement("section");
		mainSection.classList.add("main-section");

		let mainBlock = document.createElement("div");
		mainBlock.classList.add("main-block");

		let previewHolder = document.createElement("div");
		previewHolder.classList.add("bigPreviewHolder", "w3-row-padding", "w3-section");

		this._previewImg = document.createElement("img");
		this._previewImg.classList.add("bigPreview");
		this._previewImg.addEventListener('click', this.openModal);

		previewHolder.appendChild(this._previewImg);

		mainBlock.appendChild(previewHolder);

		let gallerySection = document.createElement("div");
		gallerySection.classList.add("w3-row-padding", "w3-section");

		this._thumbnailsGallery = document.createElement("div");
		this._thumbnailsGallery.classList.add("galleryList");

		gallerySection.appendChild(this._thumbnailsGallery);

		mainBlock.appendChild(gallerySection);

		mainSection.appendChild(mainBlock);
		this._container.appendChild(mainSection);

		this._modalContainer = document.createElement("div");
		this._modalContainer.classList.add('w3-modal', "styled-modal");

		let closeSpan = document.createElement("span");
		closeSpan.classList.add("w3-text-white", "w3-xxlarge", "w3-hover-text-grey",
		 "w3-container", "w3-display-topright");
		closeSpan.style.cursor = "pointer";
		closeSpan.innerHTML = "&times";
		closeSpan.addEventListener('click', this.closeModal);

		this._modalContainer.appendChild(closeSpan);

		let modalContent = document.createElement("div");
		modalContent.classList.add("w3-modal-content");

		let row = document.createElement("div");
		row.className = "w3-row";

		let imageColumn = document.createElement("div");
		imageColumn.className = "w3-display-container";

		let contentContainer = document.createElement("div");
		contentContainer.classList.add("w3-content");
		contentContainer.style.maxWidth = "1200px";
		contentContainer.style.backgroundColor = "black";

		this._modalImage = document.createElement("img");
		this._modalImage.style.maxWidth ="100%";
		this._modalImage.style.height = "80vh";

		contentContainer.appendChild(this._modalImage);

		let rightButton = document.createElement("button");
		rightButton.classList.add("w3-button", "w3-black", "w3-display-right");
		rightButton.innerHTML = "❯";
		rightButton.onclick= () => this.changeImage(1);

		let leftButton = document.createElement("button");
		leftButton.classList.add("w3-button", "w3-black", "w3-display-left");
		leftButton.innerHTML = "❮";
		leftButton.onclick= () => this.changeImage(-1);

		let infoSpan = document.createElement("img");
		infoSpan.className = "w3-display-topright w3-xlarge";
		infoSpan.src = "info32.png";
		infoSpan.style.cursor = 'pointer';
		infoSpan.style.padding = "10px";
		infoSpan.addEventListener('click', this.toggleColumn);

		imageColumn.appendChild(contentContainer);
		imageColumn.appendChild(rightButton);
		imageColumn.appendChild(leftButton);
		imageColumn.appendChild(infoSpan);
		row.appendChild(imageColumn);

		this._imageColumn = imageColumn;

		let textColumn = document.createElement("div");
		textColumn.className = "w3-hide";
		textColumn.style.color = "white";

		let p = document.createElement("p");
		p.style.color = "black";

		this._noteTitle = p;

		let hr = document.createElement("hr");
		hr.style.borderTopColor = "black";

		this._noteText = document.createElement("p");
		this._noteText.style.color = "black";

		textColumn.appendChild(this._noteTitle);
		textColumn.appendChild(hr);
		textColumn.appendChild(this._noteText);

		row.appendChild(textColumn);

		this._textColumn = textColumn;

		modalContent.appendChild(row);
		this._modalContainer.appendChild(modalContent);

		this._container.appendChild(this._modalContainer);

		container.appendChild(this._container);

		let curentRecord: ComponentFramework.EntityReference = {
			id: (<any>context).page.entityId,
			name: (<any>context).page.entityTypeName
		}
	
		this.GetAttachments(curentRecord).then( result => this.CreateGallery(result));
	
	}

	private toggleColumn():void{
		this._textColumn.classList.toggle('w3-hide');
		this._textColumn.classList.toggle('w3-rest');
		this._imageColumn.classList.toggle('w3-threequarter');
	}

	private openModal(e: Event):void {
		this._modalContainer.style.display = "block";
	}

	private closeModal(e: Event):void{
		this._modalContainer.style.display = "none";
	}

	private async GetAttachments(curentRecord: ComponentFramework.EntityReference): Promise<Attachment[]> {
		let searchQuery = "?$select=annotationid,documentbody,mimetype,notetext,subject&$filter=_objectid_value eq " + 
		curentRecord.id +
		" and  isdocument eq true and startswith(mimetype, 'image/')";
		const result = await this._context.webAPI.retrieveMultipleRecords("annotation", searchQuery);

		if(result && result.entities){
			for (let index = 0; index < result.entities.length; index++) {

				let item: Attachment = {
					id:  result.entities[index].id,
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
				newImg.className = "thumbnails w3-opacity w3-hover-opacity-off";
				newImg.src = "data:" + result[i].mimeType + ";base64," + result[i].documentBody;
				newImg.alt = count.toString();
				newImg.addEventListener('click', this.setPreviewFromThumbnail);
				this._thumbnailsGallery.appendChild(newImg);
				count++;
			}

				this._currentIndex = 0;
				this.setPreview(0);
		}
	}

	private setPreviewFromThumbnail(e: Event){
		let currentImgIndex = parseInt((e.srcElement as HTMLImageElement).alt);
		this.setPreview(currentImgIndex);
	}

	private setPreview(currentNoteNumber:number){
		if (currentNoteNumber === this._notes.length ) {currentNoteNumber = 0}
		if (currentNoteNumber < 0) {currentNoteNumber = this._notes.length-1}
		this._previewImg.src = "data:" + this._notes[currentNoteNumber].mimeType + ";base64," + this._notes[currentNoteNumber].documentBody;
		this._modalImage.src = this._previewImg.src;
		this._noteTitle.innerHTML = this._notes[currentNoteNumber].title;
		this._noteText.innerHTML = this._notes[currentNoteNumber].noteText;
		this._currentIndex = currentNoteNumber;
	}

	private changeImage(moveIndex:number){
		this.setPreview(this._currentIndex+=moveIndex);
	}
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