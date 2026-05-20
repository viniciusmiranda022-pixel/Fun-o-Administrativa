import { NavLink } from 'react-router-dom';

const tabs = [
  { to: '/dashboard', label: 'DASHBOARD' },
  { to: '/backups', label: 'BACKUPS' },
  { to: '/unpacked', label: 'UNPACKED OBJECTS' },
  { to: '/differences', label: 'DIFFERENCES' },
  { to: '/events', label: 'EVENTS' },
  { to: '/tasks', label: 'TASKS' }
];

export default function TabNav() {
  return (
    <nav className="bg-white border-b border-border-light px-2 flex">
      {tabs.map((t) => (
        <NavLink
          key={t.to}
          to={t.to}
          className={({ isActive }) =>
            `px-4 py-3 text-xs font-semibold relative ${
              isActive ? 'text-[#0078A8]' : 'text-[#333]'
            }`
          }
        >
          {({ isActive }) => (
            <>
              {t.label}
              {isActive && (
                <span className="absolute left-2 right-2 -bottom-px h-[3px] bg-[#0096D6]" />
              )}
            </>
          )}
        </NavLink>
      ))}
    </nav>
  );
}
