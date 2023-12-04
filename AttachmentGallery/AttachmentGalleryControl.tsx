import * as React from "react";
import Modal from "./components/Modal";
import { FiInfo, FiDownload, FiTrash2, FiX } from "react-icons/fi";
import { MediaControl } from "./components/MediaControl";

export interface Attachment {
  documentBody: string;
  mimeType: string;
  title?: string;
  noteText?: string;
  id: string;
  filename?: string;
}

export interface IAttachmentGallerySettings {
  showThumbnails?: boolean;
  showCaptions?: boolean;
  showPageNumbers?: boolean;
  allowDelete?: boolean;
}

export interface IAttachmentGalleryProps {
  attachments: Attachment[];
  isLoading: boolean;
  settings: IAttachmentGallerySettings;
  downloadFile: (attachment: Attachment) => void;
  deleteFile: (attachment: Attachment) => void;
}

const AttachmentGalleryControl: React.FC<IAttachmentGalleryProps> = ({
  attachments,
  settings,
  isLoading,
  downloadFile,
  deleteFile,
}) => {
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [isOpen, setIsOpen] = React.useState(false);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState(false);

  function moveSlide(slideShift: number) {
    let newIndex = currentIndex + slideShift;
    if (newIndex < 0) {
      newIndex = attachments.length - 1;
    } else if (newIndex >= attachments.length) {
      newIndex = 0;
    }
    setCurrentIndex(newIndex);
  }

  let mainContent = null;

  if (attachments.length > 0) {
    mainContent = (
      <>
        <div className="slideshow-container">
          {attachments.map((attachment, index) => (
            <div
              className={index == currentIndex ? "mySlides active" : "mySlides"}
            >
              {settings.showPageNumbers && (
                <div className="numbertext">
                  {index + 1} / {attachments.length}
                </div>
              )}
              {/* <img key={index} src={attachment.documentBody} alt={`Thumbnail ${index}`} onClick={() => setIsOpen(true)}/> */}
              <MediaControl
                attachment={attachment}
                onClick={() => setIsOpen(true)}
                key={attachment.id}
              />
              {settings.showCaptions && (
                <div className="text">Caption Text {index + 1}</div>
              )}
            </div>
          ))}
          <a className="prev" onClick={() => moveSlide(-1)}>
            &#10094;
          </a>
          <a className="next" onClick={() => moveSlide(1)}>
            &#10095;
          </a>
        </div>
        <div className="thumbnails">
          {settings.showThumbnails &&
            attachments.map((attachment, index) => (
              <img
                key={index}
                className={
                  index == currentIndex ? "thumbnail active" : "thumbnail"
                }
                src={attachment.documentBody}
                onClick={() => setCurrentIndex(index)}
                alt={`Thumbnail ${index}`}
              />
            ))}
        </div>
        <Modal openModal={isOpen} closeModal={() => setIsOpen(false)}>
          <div>
            <header className="custom-dialog__header">
              <div>
                {currentIndex + 1}/{attachments.length}
              </div>
              <div>
                <button className="dialog-button">
                  <FiInfo />
                </button>
                <button
                  className="dialog-button"
                  aria-label="Download File"
                  onClick={() => downloadFile(attachments[currentIndex])}
                >
                  <FiDownload />
                </button>
                <button
                  className="dialog-button"
                  aria-label="Delete File"
                  onClick={() => setIsDeleteDialogOpen(true)}
                >
                  <FiTrash2 />
                </button>
                <button
                  className="dialog-button"
                  aria-label="Close modal"
                  onClick={() => setIsOpen(false)}
                >
                  <FiX />
                </button>
              </div>
            </header>
            <div className="custom-dialog__body">
              <img src={attachments[currentIndex].documentBody} />
              <a className="prev" onClick={() => moveSlide(-1)}>
                &#10094;
              </a>
              <a className="next" onClick={() => moveSlide(1)}>
                &#10095;
              </a>
            </div>
          </div>
        </Modal>
        <Modal
          openModal={isDeleteDialogOpen}
          closeModal={() => setIsDeleteDialogOpen(false)}
        >
          <div className="delete-dialog">
            <p>Are you sure you want to delete this file?</p>
            <div className="delete-dialog__button-container">
              <button
                className="delete-dialog__button delete-dialog__button-confirm"
                onClick={() => {
                  setIsDeleteDialogOpen(false);
                  deleteFile(attachments[currentIndex]);
                }}
              >
                Delete
              </button>
              <button
                className="delete-dialog__button"
                onClick={() => setIsDeleteDialogOpen(false)}
              >
                Cancel
              </button>
            </div>
          </div>
        </Modal>
      </>
    );
  } else {
    mainContent = <div>No records found</div>;
  }

  return (
    <div className="main-container">
      {isLoading ? <div className="loader"></div> : mainContent}
    </div>
  );
};

export default AttachmentGalleryControl;
