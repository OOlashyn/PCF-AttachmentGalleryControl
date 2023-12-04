import * as React from "react";

import { Attachment } from "../AttachmentGalleryControl";
import { PDFViewer } from "./PDFViewer";

interface IMediaControlProps {
  attachment: Attachment;
}

export const MediaControl = ({ attachment, onClick }: { attachment: Attachment, onClick: () => void }) => {

  switch (attachment.mimeType) {
    case "application/pdf":
      return (
        <div className="pdf-container">
          <PDFViewer attachment={attachment}/>
        </div>
      );
    case "image/png":
    case "image/jpeg":
    case "image/gif":
      return <img src={attachment.documentBody} alt={attachment.filename} onClick={onClick}/>;
    default:
      return <div>Unsupported file type</div>;
  }
};
