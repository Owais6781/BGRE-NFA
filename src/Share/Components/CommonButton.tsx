import * as React from 'react';
import "../styles/global.scss";

interface ICommonButtonProps {
  text: string;
  onClick?: () => void;
  type?: 'button' | 'submit' | 'reset';
  variant?: 'primary' | 'danger' | 'secondary';
  disabled?: boolean;
}

const CommonButton: React.FC<ICommonButtonProps> = ({
  text, onClick, type = 'button', variant = 'primary', disabled = false,
}) => {
  const cls: Record<string, string> = {
    primary:   'btnPrimary',
    danger:    'btnDanger',
    secondary: 'btnSecondary',
  };

  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className={`btn ${cls[variant]} ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
    >
      {text}
    </button>
  );
};

export default CommonButton;
