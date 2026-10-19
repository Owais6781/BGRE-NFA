import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { NavLink, useLocation, useNavigate } from 'react-router-dom';
import {
  Home, FilePlus, Inbox, Eye, CheckCheck,
  List, Settings2, BarChart2, Clock, TrendingUp,
  CalendarDays, FileText, ChevronLeft, ChevronRight, Headphones
} from 'lucide-react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { NFAService } from '../../../../Share/Services/NFAService';

import "../../../../Share/styles/global.scss";

type NavItem =
  | { kind: 'section'; label: string }
  | { kind: 'link'; to: string; label: string; icon: React.ElementType; badge?: string; end?: boolean };

const NAV: NavItem[] = [
  { kind: 'section', label: 'Navigation' },
  { kind: 'link', to: '/',                    label: 'Home',                      icon: Home,          end: true },
  { kind: 'link', to: '/new-nfa',             label: 'New NFA',                   icon: FilePlus },
  { kind: 'link', to: '/inbox',               label: 'My Inbox',                  icon: Inbox },
  { kind: 'link', to: '/approver-detail',     label: "Approver's Detail View",    icon: Eye },
  { kind: 'link', to: '/approved-rejected',   label: 'Approved / Rejected by Me', icon: CheckCheck },
  { kind: 'section', label: 'Admin Section' },
  { kind: 'link', to: '/detail-view',         label: 'Detail View',               icon: List },
  { kind: 'link', to: '/masters',             label: 'Masters',                   icon: Settings2 },
  { kind: 'section', label: 'Reports' },
  { kind: 'link', to: '/function-wise-status',   label: 'Function Wise Status',      icon: BarChart2 },
  { kind: 'link', to: '/approver-wise-pendency', label: 'Approver Wise Pendency',    icon: Clock },
  { kind: 'link', to: '/ageing-report',          label: 'Ageing Report',             icon: TrendingUp },
  { kind: 'link', to: '/day-wise-ageing',        label: 'Day Wise Ageing',           icon: CalendarDays },
  { kind: 'link', to: '/generic-report',         label: 'Generic Report',            icon: FileText },
];

interface ISidebarProps {
  context?: WebPartContext;
  sp?: SPFI;
}

const Sidebar: React.FC<ISidebarProps> = ({ context, sp }) => {
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [inboxCount, setInboxCount] = useState<number | null>(null);
  const [notification, setNotification] = useState<{ show: boolean; subject: string; requestNo: string; id: number } | null>(null);
  
  const location = useLocation();
  const navigate = useNavigate();
  const isInitialLoad = useRef(true);
  const previousItems = useRef<any[]>([]);

  useEffect(() => {
    if (!context || !sp) return;
    
    const fetchCount = async () => {
      try {
        const service = new NFAService(context);
        const data = await service.getInboxData();
        setInboxCount(data.length);

        if (!isInitialLoad.current) {
          const newItems = data.filter(newItem => !previousItems.current.some(oldItem => oldItem.id === newItem.id));
          if (newItems.length > 0) {
            const item = newItems[0];
            setNotification({
              show: true,
              subject: item.project || "New NFA Assigned",
              requestNo: item.requestNo || `NFA #${item.id}`,
              id: item.id
            });
            
            setTimeout(() => {
              setNotification(null);
            }, 6000);
          }
        } else {
          isInitialLoad.current = false;
        }

        previousItems.current = data;
      } catch (err) {
        console.error("Sidebar: Error fetching inbox count", err);
      }
    };
    
    void fetchCount();

    // Poll every 30 seconds for real-time updates
    const interval = setInterval(fetchCount, 30000);
    return () => clearInterval(interval);
  }, [context, sp, location.pathname]);

  const handleNotificationClick = () => {
    if (notification) {
      navigate(`/approver-page/${notification.id}`);
      setNotification(null);
    }
  };

  const linkClass = ({ isActive }: { isActive: boolean }) =>
    `sidebarLink${isActive ? ' active' : ''}`;

  return (
    <div className={`sidebarWrapper${isCollapsed ? ' sidebarCollapsed' : ''}`}>

      {/* Header with toggle */}
      <div className={`sidebarHeader ${isCollapsed ? 'justify-center !px-0' : 'justify-end !px-3'} !h-[50px]`}>
        <button
          className="sidebarToggleBtn"
          onClick={() => setIsCollapsed((c) => !c)}
          title={isCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}
        >
          {isCollapsed
            ? <ChevronRight size={14} strokeWidth={2.5} />
            : <ChevronLeft  size={14} strokeWidth={2.5} />
          }
        </button>
      </div>

      {/* Scrollable nav */}
      <div className="sidebar flex flex-col justify-between">
        <div>
          {NAV.map((item, i) => {
            if (item.kind === 'section') {
              return (
                <div key={i} className="sidebarSectionTitle">
                  {item.label}
                </div>
              );
            }

            const Icon = item.icon;
            return (
              <NavLink
                key={item.to}
                to={item.to}
                end={item.end}
                className={linkClass}
                title={isCollapsed ? item.label : undefined}
              >
                <span className="sidebarLinkIcon">
                  <Icon size={16} strokeWidth={1.75} />
                </span>
                <span className="sidebarLinkLabel">{item.label}</span>
                {((item.to === '/inbox' ? (inboxCount && inboxCount > 0 ? String(inboxCount) : undefined) : item.badge)) && (
                  <span className="sidebarBadge">
                    {item.to === '/inbox' ? inboxCount : item.badge}
                  </span>
                )}
              </NavLink>
            );
          })}
        </div>

        {/* Support Box */}
        <div className="mt-8">
          {/* Divider matching section titles */}
          <div style={{ borderTop: '1px solid var(--sidebar-border, #E2E8F0)' }}></div>
          
          {!isCollapsed && (
            <div className="m-4 mt-6 p-4 rounded-xl flex flex-col gap-3" style={{ background: 'var(--primary-burgundy-light, #F7EBEF)' }}>
              <div className="flex items-center gap-3">
                <div className="p-2 rounded-full" style={{ background: '#ffffff' }}>
                  <Headphones size={24} color="#4A081E" strokeWidth={1.5} />
                </div>
                <div className="flex flex-col">
                  <span className="text-[#4A081E] font-semibold text-sm">Need Help?</span>
                </div>
              </div>
              <p className="text-xs text-slate-600 leading-snug">
                Our support team is available to assist you.
              </p>
              <button className="w-full py-2 rounded-lg font-semibold text-sm transition-colors mt-1" style={{ background: 'var(--primary-burgundy, #4A081E)', color: '#ffffff' }}>
                Contact Support
              </button>
            </div>
          )}
          {isCollapsed && (
            <div className="mx-auto mt-6 mb-4 p-2 rounded-full cursor-pointer transition-colors flex items-center justify-center w-[36px] h-[36px]" title="Contact Support" style={{ background: 'var(--primary-burgundy-light, #F7EBEF)' }}>
               <Headphones size={20} color="#4A081E" strokeWidth={1.5} />
            </div>
          )}
        </div>
      </div>

      {/* ── Toast Notification ── */}
      {notification && (
        <div 
          className="fixed bottom-6 right-6 z-[9999] bg-white/95 backdrop-blur-md rounded-2xl border border-slate-200 shadow-2xl p-4 w-[320px] animate-slide-in flex items-start gap-3.5 transition-all hover:-translate-y-1 duration-350 cursor-pointer" 
          onClick={handleNotificationClick}
        >
          <div className="p-2 bg-blue-50 text-blue-600 rounded-xl border border-blue-100 shrink-0">
            <Inbox size={20} />
          </div>
          <div className="flex-1 min-w-0">
            <div className="flex items-center justify-between gap-1 mb-0.5">
              <span className="text-[10px] font-extrabold text-blue-600 uppercase tracking-wider">New NFA Assigned</span>
              <button 
                className="text-slate-400 hover:text-slate-655 font-extrabold text-xs p-1"
                onClick={(e) => {
                  e.stopPropagation();
                  setNotification(null);
                }}
              >
                ✕
              </button>
            </div>
            <p className="text-xs font-extrabold text-slate-800 truncate mb-1" title={notification.subject}>{notification.subject}</p>
            <p className="text-[10.5px] text-slate-500 font-bold leading-normal truncate">{notification.requestNo}</p>
          </div>
        </div>
      )}

      <style>{`
        @keyframes slideInUp {
          from {
            transform: translateY(100%) scale(0.95);
            opacity: 0;
          }
          to {
            transform: translateY(0) scale(1);
            opacity: 1;
          }
        }
        .animate-slide-in {
          animation: slideInUp 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        }
      `}</style>

    </div>
  );
};

export default Sidebar;
