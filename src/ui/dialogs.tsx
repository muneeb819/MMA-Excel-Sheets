/** Small modal wrapper used by the dialogs in App. */

import type { ReactNode } from 'react';

export function Modal({
  title,
  children,
  onClose,
}: {
  title: string;
  children: ReactNode;
  onClose: () => void;
}) {
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" role="dialog" aria-label={title}>
        <header>{title}</header>
        <div className="body">{children}</div>
      </div>
    </div>
  );
}

export default Modal;
