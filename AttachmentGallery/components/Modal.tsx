// Modal as a separate component
import * as React from "react";

interface IModalProps {
  openModal: boolean;
  closeModal: () => void;
  children: React.ReactNode;
}

const Modal: React.FC<IModalProps> = ({ openModal, closeModal, children }) => {
  const ref = React.useRef();

  React.useEffect(() => {
    if (openModal) {
      // @ts-ignore
      ref.current?.showModal();
    } else {
      // @ts-ignore
      ref.current?.close();
    }
  }, [openModal]);

  return (
    // @ts-ignore
    <dialog ref={ref} onCancel={closeModal} className="custom-dialog">
      {children}
    </dialog>
  );
};

export default Modal;
