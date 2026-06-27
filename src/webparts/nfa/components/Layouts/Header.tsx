import React, { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../../../../Share/styles/global.scss";

interface HeaderProps {
  context: WebPartContext;
}

const logo = require('../../assets/Leelalogo-black.svg');

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
    <header className="appHeader flex items-center justify-between !bg-transparent !shadow-none !border-none !h-[50px]" style={{ paddingRight: '32px' }}>
      {/* Left — Title with Leela logo */}
      <div className="flex items-center gap-3.5" style={{ marginLeft: '38px' }}>
        <img 
          src={logo} 
          alt="Leela Logo" 
          style={{ height: '32px', display: 'block' }} 
        />
        <div className="h-5 w-[1px] bg-slate-300"></div>
        <span className="headerTitle uppercase tracking-widest text-[14px] !text-slate-800 opacity-95 font-bold">Note for Approval (NFA)</span>
      </div>

      {/* Right — user info aligned with form end */}
      <div className="flex items-center gap-6">
        <div className="h-6 w-[1px] bg-slate-300"></div>

        <div className="flex items-center gap-3">
          <span className="text-[13px] font-medium !text-slate-800 tracking-tight">{userName}</span>

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
