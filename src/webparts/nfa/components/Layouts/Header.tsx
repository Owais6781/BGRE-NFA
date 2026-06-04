import React, { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../../../../Share/styles/global.scss";

interface HeaderProps {
  context: WebPartContext;
}

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
    <header className="appHeader flex items-center justify-between" style={{ paddingRight: '32px' }}>
      {/* Left — Title with specific 50px margin */}
      <div className="flex items-center" style={{ marginLeft: '38px' }}>
        <span className="headerTitle uppercase tracking-widest text-[14px] text-white opacity-95">Note for Approval (NFA)</span>
      </div>

      {/* Right — user info aligned with form end */}
      <div className="flex items-center gap-6">
        <div className="h-6 w-[1px] bg-white/20"></div>

        <div className="flex items-center gap-3">
          <span className="text-[13px] font-medium text-white/90 tracking-tight">{userName}</span>

          <div className="headerAvatar overflow-hidden border-white/20">
            {userPhoto ? (
              <img
                src={userPhoto}
                alt={userName}
                className="w-full h-full object-cover"
                onError={() => setUserPhoto("")}
              />
            ) : (
              <span className="text-[11px] font-bold text-white/80">{userInitials}</span>
            )}
          </div>
        </div>
      </div>
    </header>
  );
};

export default Header;
