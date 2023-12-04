import { createElement } from "react";
import { createRoot, Root } from "react-dom/client";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import AttachmentGalleryControl, {
  Attachment,
  IAttachmentGalleryProps,
} from "./AttachmentGalleryControl";
import { samplePdf } from "./sample/samplePdf";
import { saveAs } from 'file-saver';

export class AttachmentGallery
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _notifyOutputChanged: () => void;
  private _root: Root;
  private _props: IAttachmentGalleryProps;
  private _context: ComponentFramework.Context<IInputs>;

  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._notifyOutputChanged = notifyOutputChanged;
    this._context = context;

    this.downloadFile = this.downloadFile.bind(this);
    this.deleteFile = this.deleteFile.bind(this);

    this._root = createRoot(container);

    let curentRecord: ComponentFramework.EntityReference = {
      id: (<any>context).page.entityId,
      name: (<any>context).page.entityTypeName,
    };

    this._props = {
      attachments: [],
      settings: {
        showCaptions: true,
        showPageNumbers: true,
        showThumbnails: true,
        allowDelete: true,
      },
      downloadFile: this.downloadFile,
      deleteFile: this.deleteFile,
      isLoading: true,
    };

    setTimeout(async () => {
      let attachments = await this.GetAttachments(curentRecord);

      this._props.attachments = attachments;
      this._props.isLoading = false;

      this._root.render(createElement(AttachmentGalleryControl, this._props));

    }, 1000);
  }

  private async GetAttachments(
    curentRecord: ComponentFramework.EntityReference
  ): Promise<Attachment[]> {
    let attachments: Attachment[] = [];
    let imagesReq = await fetch("https://picsum.photos/v2/list");
    await imagesReq.json().then((images: any) => {
      images.forEach((image: any) => {
        attachments.push({
          documentBody: image.download_url,
          mimeType: "image/jpeg",
          title: image.author,
          noteText: image.author,
          id: image.id,
          filename: image.author,
        });
      });
    });

    attachments.push({
      documentBody: samplePdf,
      mimeType: "application/pdf",
      title: "PDF",
      noteText: "PDF",
      id: "pdf",
      filename: "PDF",
    });

    return attachments;
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

  private downloadFile(attachment: Attachment): void {
      let blob = this.b64toBlob(attachment.documentBody, attachment.mimeType, 512);
      saveAs(blob, attachment.filename);
  }

  private deleteFile(attachment: Attachment): void {
    this._context.webAPI.deleteRecord("annotation", attachment.id).then(
      function success(result) {
        console.log(result);
      },
      function(error) {
        console.log(error.message);
      }
    );
  }


  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._root.render(createElement(AttachmentGalleryControl, this._props));
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
    this._root.unmount();
  }
}
