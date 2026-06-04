import * as React from "react";
import { useState, useEffect } from "react";
import { CheckCircle, XCircle, AlertCircle, X } from "lucide-react";
import "../styles/global.scss";

export type ToastType = "success" | "error" | "warning";

export interface IToastMessage {
  id: string;
  type: ToastType;
  title: string;
  message: string;
}

interface IToastNotificationProps {
  toasts: IToastMessage[];
  onRemove: (id: string) => void;
}

const ToastIcons: Record<ToastType, React.ReactNode> = {
  success: <CheckCircle size={18} />,
  error: <XCircle size={18} />,
  warning: <AlertCircle size={18} />,
};

const ToastNotification: React.FC<IToastNotificationProps> = ({
  toasts,
  onRemove,
}) => {
  return (
    <div className="toastContainer">
      {toasts.map((toast) => (
        <ToastItem key={toast.id} toast={toast} onRemove={onRemove} />
      ))}
    </div>
  );
};

const ToastItem: React.FC<{
  toast: IToastMessage;
  onRemove: (id: string) => void;
}> = ({ toast, onRemove }) => {
  useEffect(() => {
    const timer = setTimeout(() => {
      onRemove(toast.id);
    }, 5000);
    return () => clearTimeout(timer);
  }, [toast.id, onRemove]);

  return (
    <div className={`toast toast--${toast.type}`}>
      <div className="toast__icon">{ToastIcons[toast.type]}</div>
      <div className="toast__content">
        <p className="toast__title">{toast.title}</p>
        <p className="toast__message">{toast.message}</p>
      </div>
      <button className="toast__close" onClick={() => onRemove(toast.id)}>
        <X size={14} />
      </button>
    </div>
  );
};

export default ToastNotification;