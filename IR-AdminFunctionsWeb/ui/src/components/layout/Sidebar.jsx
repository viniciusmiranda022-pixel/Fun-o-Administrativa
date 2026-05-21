import { NavLink } from 'react-router-dom';
import { Home, Users, RotateCw, Settings, HelpCircle } from 'lucide-react';

const items = [
  { to: '/dashboard', icon: Home,        label: 'Home' },
  { to: '/tenants',   icon: Users,       label: 'Tenants' },
  { to: '/backups',   icon: RotateCw,    label: 'Recover',  alwaysAccent: true },
  { to: '/setup',     icon: Settings,    label: 'Settings' },
  { to: '/support',   icon: HelpCircle,  label: 'Support' },
];

export default function Sidebar() {
  return (
    <aside className="w-[68px] bg-[#1F1F1F] flex flex-col items-center py-1 shrink-0">
      {items.map(({ to, icon: Icon, label, alwaysAccent }) => (
        <NavLink
          key={to}
          to={to}
          className={({ isActive }) =>
            `relative w-full flex flex-col items-center py-3 cursor-pointer select-none
             ${isActive ? 'bg-white/10' : 'hover:bg-white/5'}`
          }
        >
          {({ isActive }) => (
            <>
              {(isActive || alwaysAccent) && (
                <span className="absolute left-0 top-0 bottom-0 w-[3px] bg-[#0096D6] rounded-r-sm" />
              )}
              <Icon size={18} className={isActive ? 'text-white' : 'text-[#aaa]'} />
              <span className={`text-[10px] mt-1 leading-none ${isActive ? 'text-white' : 'text-[#999]'}`}>
                {label}
              </span>
            </>
          )}
        </NavLink>
      ))}
    </aside>
  );
}
