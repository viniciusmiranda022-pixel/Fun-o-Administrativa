import { NavLink } from 'react-router-dom';

const TABS = [
  { to: '/dashboard', label: 'DASHBOARD' },
  { to: '/backups',   label: 'BACKUPS' },
  { to: '/unpacked',  label: 'UNPACKED OBJECTS' },
  { to: '/differences', label: 'DIFFERENCES' },
  { to: '/events',    label: 'EVENTS' },
  { to: '/tasks',     label: 'TASKS' },
];

export default function TabNav() {
  return (
    <nav className="bg-white border-b border-border-light px-4 flex flex-shrink-0">
      {TABS.map((t) => (
        <NavLink
          key={t.to}
          to={t.to}
          className={({ isActive }) =>
            `px-4 py-3 text-[11px] font-semibold relative tracking-wider whitespace-nowrap transition-colors
             ${isActive ? 'text-[#0078A8]' : 'text-[#555] hover:text-[#222]'}`
          }
        >
          {({ isActive }) => (
            <>
              {t.label}
              {isActive && (
                <span className="absolute left-0 right-0 -bottom-px h-[3px] bg-[#0096D6]" />
              )}
            </>
          )}
        </NavLink>
      ))}
    </nav>
  );
}
