import * as React from 'react';
import { useState } from 'react';
import { NavLink } from 'react-router-dom';
import {
  Home, FilePlus, Inbox, Eye, CheckCheck,
  List, Settings2, BarChart2, Clock, TrendingUp,
  CalendarDays, FileText, ChevronLeft, ChevronRight, Headphones
} from 'lucide-react';
const logo = require('../../assets/Leelalogo-black.svg');
import "../../../../Share/styles/global.scss";

type NavItem =
  | { kind: 'section'; label: string }
  | { kind: 'link'; to: string; label: string; icon: React.ElementType; badge?: string; end?: boolean };

const NAV: NavItem[] = [
  { kind: 'section', label: 'Navigation' },
  { kind: 'link', to: '/',                    label: 'Home',                      icon: Home,          end: true },
  { kind: 'link', to: '/new-nfa',             label: 'New NFA',                   icon: FilePlus },
  { kind: 'link', to: '/inbox',               label: 'My Inbox',                  icon: Inbox,         badge: '4' },
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

const Sidebar = () => {
  const [isCollapsed, setIsCollapsed] = useState(false);

  const linkClass = ({ isActive }: { isActive: boolean }) =>
    `sidebarLink${isActive ? ' active' : ''}`;

  return (
    <div className={`sidebarWrapper${isCollapsed ? ' sidebarCollapsed' : ''}`}>

      {/* Header with logo and toggle */}
      <div className={`sidebarHeader ${isCollapsed ? 'flex-col gap-3 py-3 h-auto' : 'h-[80px]'}`}>
        <div className="flex items-center justify-center">
          <img 
            src={logo} 
            alt="Leela Logo" 
            style={{ 
              height: isCollapsed ? '28px' : '60px', 
              filter: 'brightness(0) invert(1)',
              transition: 'all 0.2s ease',
              display: 'block'
            }} 
          />
        </div>
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
                {item.badge && (
                  <span className="sidebarBadge">{item.badge}</span>
                )}
              </NavLink>
            );
          })}
        </div>

        {/* Support Box */}
        <div className="mt-8">
          {/* Divider matching section titles */}
          <div style={{ borderTop: '1px solid rgba(255, 255, 255, 0.08)' }}></div>
          
          {!isCollapsed && (
            <div className="m-4 mt-6 p-4 rounded-xl flex flex-col gap-3" style={{ background: 'rgba(255, 255, 255, 0.05)' }}>
              <div className="flex items-center gap-3">
                <div className="p-2 rounded-full" style={{ background: 'rgba(255, 255, 255, 0.1)' }}>
                  <Headphones size={24} color="#ffffff" strokeWidth={1.5} />
                </div>
                <div className="flex flex-col">
                  <span className="text-white font-semibold text-sm">Need Help?</span>
                </div>
              </div>
              <p className="text-xs text-slate-300 leading-snug">
                Our support team is available to assist you.
              </p>
              <button className="w-full py-2 rounded-lg font-semibold text-sm transition-colors mt-1" style={{ background: '#AF8941', color: '#ffffff' }}>
                Contact Support
              </button>
            </div>
          )}
          {isCollapsed && (
            <div className="mx-auto mt-6 mb-4 p-2 rounded-full cursor-pointer hover:bg-white/10 transition-colors flex items-center justify-center w-[36px] h-[36px]" title="Contact Support" style={{ background: 'rgba(255, 255, 255, 0.05)' }}>
               <Headphones size={20} color="#ffffff" strokeWidth={1.5} />
            </div>
          )}
        </div>
      </div>

    </div>
  );
};

export default Sidebar;
