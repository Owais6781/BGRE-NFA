import * as React from 'react';
import '../styles/global.scss';

export interface PageTitleProps {
  title: string;
  subtitle?: string;
  action?: React.ReactNode;
  className?: string;
}

export const PageTitle: React.FC<PageTitleProps> = ({
  title,
  subtitle,
  action,
  className = ''
}) => {
  return (
    <div 
      className={`nfaTitleSection flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-5 ${className}`}
      style={{ display: 'flex', marginBottom: '20px' }}
    >
      <div className="flex flex-col">
        <h1 
          className="nfaPageTitle"
          style={{
            fontFamily: 'Georgia, Cambria, "Times New Roman", Times, serif',
            fontSize: '26px',
            fontWeight: 700,
            color: '#4A081E',
            margin: '0 0 4px 0',
            lineHeight: 1.25,
            letterSpacing: '-0.2px'
          }}
        >
          {title}
        </h1>
        {subtitle && (
          <p 
            className="nfaPageSubtitle"
            style={{
              fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif',
              fontSize: '13.5px',
              fontWeight: 400,
              color: '#71717A',
              margin: 0,
              lineHeight: 1.4,
              letterSpacing: '0.1px'
            }}
          >
            {subtitle}
          </p>
        )}
      </div>
      {action && (
        <div className="nfaPageAction flex items-center gap-3 shrink-0">
          {action}
        </div>
      )}
    </div>
  );
};

export default PageTitle;
