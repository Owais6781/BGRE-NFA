import * as React from "react";
import { CheckCircle, XCircle, AlertTriangle, Info, X } from "lucide-react";
import "../styles/global.scss";

// ── Message Types ──
export type MessageType = "success" | "error" | "warning" | "info";

export interface IMessagePopupProps {
  isOpen: boolean;
  type: MessageType;
  title: string;
  messages: string[];
  onClose: () => void;
}

const IconMap: Record<MessageType, React.ReactNode> = {
  success: <CheckCircle size={28} />,
  error: <XCircle size={28} />,
  warning: <AlertTriangle size={28} />,
  info: <Info size={28} />,
};

const MessagePopup: React.FC<IMessagePopupProps> = ({
  isOpen,
  type,
  title,
  messages,
  onClose,
}) => {
  if (!isOpen) return null;

  return (
    <div className="msgPopupOverlay">
      <div className={`msgPopupBox msgPopupBox--${type}`}>

        {/* ── Close Button ── */}
        <button className="msgPopupBox__close" onClick={onClose}>
          <X size={16} />
        </button>

        {/* ── Icon ── */}
        <div className={`msgPopupBox__icon msgPopupBox__icon--${type}`}>
          {IconMap[type]}
        </div>

        {/* ── Title ── */}
        <h3 className="msgPopupBox__title">{title}</h3>

        {/* ── Messages ── */}
        <div className="msgPopupBox__body">
          {messages.length === 1 ? (
            <p className="msgPopupBox__singleMsg">{messages[0]}</p>
          ) : (
            <ul className="msgPopupBox__list">
              {messages.map((msg, idx) => (
                <li key={idx}>{msg}</li>
              ))}
            </ul>
          )}
        </div>

        {/* ── OK Button ── */}
        <button
          className={`msgPopupBox__btn msgPopupBox__btn--${type}`}
          onClick={onClose}
        >
          OK
        </button>

      </div>
    </div>
  );
};

export default MessagePopup;