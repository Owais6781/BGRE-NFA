import React, { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../../../../Share/styles/global.scss";

interface HeaderProps {
  context: WebPartContext;
}

const logo = require('../../assets/COMM-IT-HD-Logo-2.png');

const Header: React.FC<HeaderProps> = ({ context }) => {
  const [userName, setUserName] = useState<string>("");
  const [userInitials, setUserInitials] = useState<string>("");
  const [userPhoto, setUserPhoto] = useState<string>("");

  useEffect(() => {
    if (context?.pageContext?.user) {
      const displayName = context.pageContext.user.displayName || "";
      const email = context.pageContext.user.email || "";

      setUserName(displayName);

      const parts = displayName.trim().split(" ");
      setUserInitials(
        parts.length >= 2
          ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
          : displayName.slice(0, 2).toUpperCase()
      );

      if (email) {
        const photoUrl = `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;
        setUserPhoto(photoUrl);
      }
    }
  }, [context]);

  return (
    <header 
      className="appHeader sticky top-0 z-[1000] flex items-center justify-between !border-b !rounded-none !h-[64px] px-4 sm:px-6 w-full shrink-0 shadow-sm"
      style={{
        background: 'var(--header-background, #F7EBEF)',
        borderColor: '#EFDDE3',
        position: 'sticky',
        top: 0,
        zIndex: 1000
      }}
    >
      {/* Left — COMM-IT Logo + Application Title */}
      <div className="flex items-center gap-3.5">
        <img 
          src={logo} 
          alt="COMM-IT Logo" 
          style={{ height: '38px', width: 'auto', objectFit: 'contain', display: 'block' }} 
        />
        <div className="h-6 w-[1px] bg-slate-300"></div>
        <span className="headerTitle truncate">
          Note for Approval
        </span>
      </div>

      {/* Right — user info aligned with form end */}
      <div className="flex items-center gap-4 sm:gap-6 shrink-0">
        <div className="headerUserChip">
          <div className="headerAvatar overflow-hidden">
            {userPhoto ? (
              <img
                src={userPhoto}
                alt={userName}
                className="w-full h-full object-cover"
                onError={() => setUserPhoto("")}
              />
            ) : (
              <span className="text-[11px] font-bold">{userInitials}</span>
            )}
          </div>
          <span className="headerUserName hidden sm:inline">{userName}</span>
        </div>
      </div>
    </header>
  );
};

export default Header;
