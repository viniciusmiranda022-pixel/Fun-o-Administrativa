import { Home, Users, RotateCcw, Settings, HelpCircle, ChevronRight } from 'lucide-react';
import { useLocation, useNavigate } from 'react-router-dom';

const MENU = [
  { id: 'home',     Icon: Home,       label: 'Home',    to: '/dashboard' },
  { id: 'tenants',  Icon: Users,      label: 'Tenants', to: '/tenants'  },
  { id: 'recover',  Icon: RotateCcw,  label: 'Recover', to: '/dashboard' },
  { id: 'settings', Icon: Settings,   label: 'Settings',to: '/setup'    },
  { id: 'support',  Icon: HelpCircle, label: 'Support', to: null        },
];

const RECOVER_PREFIXES = [
  '/dashboard', '/backups', '/unpacked', '/differences',
  '/events', '/tasks', '/manage-', '/unpack',
];

function getActive(pathname) {
  if (pathname === '/tenants' || pathname.startsWith('/tenants')) return 'tenants';
  if (pathname === '/setup') return 'settings';
  if (RECOVER_PREFIXES.some((p) => pathname.startsWith(p))) return 'recover';
  return 'home';
}

export default function Sidebar() {
  const { pathname } = useLocation();
  const navigate = useNavigate();
  const activeId = getActive(pathname);

  return (
    <div className="flex h-full flex-shrink-0">
      {/* Icon strip */}
      <div className="w-[64px] bg-[#292929] flex flex-col items-center py-2 gap-px">
        {MENU.map(({ id, Icon, label, to }) => {
          const active = id === activeId;
          return (
            <button
              key={id}
              onClick={() => to && navigate(to)}
              title={label}
              className={`relative w-full flex flex-col items-center justify-center py-3 gap-1 transition-colors
                ${active ? 'text-white bg-[#3C3C3C]' : 'text-[#888] hover:text-white hover:bg-[#333]'}`}
            >
              {active && (
                <span className="absolute left-0 top-0 bottom-0 w-[3px] bg-[#0096D6]" />
              )}
              <Icon size={17} />
              <span className="text-[9px] leading-tight">{label}</span>
            </button>
          );
        })}
      </div>

      {/* Recover submenu — only shown when in the Recover section */}
      {activeId === 'recover' && (
        <div className="w-[150px] bg-[#383838] flex flex-col flex-shrink-0">
          <div className="px-3 pt-4 pb-2 text-[9px] text-[#777] uppercase tracking-widest font-semibold">
            Recover
          </div>
          <button
            onClick={() => navigate('/dashboard')}
            className="flex items-center justify-between px-3 py-2.5 text-[11px] text-white bg-[#2B2B2B] border-l-[3px] border-[#0096D6]"
          >
            <span>Microsoft Entra ID</span>
            <ChevronRight size={11} className="text-[#555]" />
          </button>
        </div>
      )}
    </div>
  );
}
