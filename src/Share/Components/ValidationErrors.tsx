import * as React from "react";
import { AlertCircle } from "lucide-react";
import { IValidationError } from "../Services/NFAService";
import "../styles/global.scss";

interface IValidationErrorsProps {
  errors: IValidationError[];
  onClose?: () => void;
}

const ValidationErrors: React.FC<IValidationErrorsProps> = ({
  errors,
  onClose,
}) => {
  if (errors.length === 0) return null;

  return (
    <div className="validationBanner">
      <div className="validationBanner__header">
        <div className="validationBanner__title">
          <AlertCircle size={16} />
          <span>
            Please fill the following {errors.length} field{errors.length > 1 ? "s" : ""} value:
          </span>
        </div>
        {onClose && (
        <button className="validationBanner__close" onClick={onClose}>
          ✕
        </button>
         )}
      </div>
      <ul className="validationBanner__list">
        {errors.map((err, idx) => (
          <li key={idx}>{err.message}</li>
        ))}
      </ul>
    </div>
  );
};

export default ValidationErrors;