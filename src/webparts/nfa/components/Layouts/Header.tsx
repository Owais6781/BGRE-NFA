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
    <header className="appHeader flex items-center justify-between px-4 h-[54px]">
      {/* Left — logo + title */}
      <div className="flex items-center">
        <div className="flex flex-col items-center leading-none">
          <span className="headerLogoScript">The Leela</span>
          <span className="headerLogoSub">PALACES · HOTELS · RESORTS</span>
        </div>

        <div className="headerDivider" />

        <span className="headerTitle">NFA — Note for Approval</span>
      </div>

      {/* Right — user info */}
      <div className="flex items-center gap-3">
        <span className="headerUserName">{userName}</span>

        <div className="headerAvatar overflow-hidden">
          {userPhoto ? (
            <img
              src={userPhoto}
              alt={userName}
              className="w-full h-full object-cover"
              onError={() => setUserPhoto("")}
            />
          ) : (
            <span className="text-[12px] font-semibold">{userInitials}</span>
          )}
        </div>
      </div>
    </header>
  );
};

export default Header;
