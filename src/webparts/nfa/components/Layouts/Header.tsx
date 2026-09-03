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

      setUserName(displayName.toUpperCase());

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
      className="appHeader flex items-center justify-between !shadow-none !border-b !rounded-none !h-[64px] px-4 sm:px-6 w-full shrink-0"
      style={{
        background: 'var(--header-background, #FAF9F6)',
        borderColor: 'var(--border-color, #E8E6E1)'
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
        <span className="headerTitle uppercase tracking-widest text-[14px] !text-slate-800 opacity-95 font-bold truncate">
          Note for Approval (NFA)
        </span>
      </div>

      {/* Right — user info aligned with form end */}
      <div className="flex items-center gap-4 sm:gap-6 shrink-0">
        <div className="flex items-center gap-3">
          <span className="text-[13px] font-medium !text-slate-800 tracking-tight hidden sm:inline">{userName}</span>

          <div className="headerAvatar overflow-hidden !border-slate-300 !bg-slate-100 !text-slate-700">
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
        </div>
      </div>
    </header>
  );
};

export default Header;
