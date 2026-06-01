import { useLocation, useNavigate } from 'react-router-dom';
import { Home, Building2, RotateCw, Settings, HelpCircle } from 'lucide-react';

const RECOVER_PATHS = ['/dashboard', '/backups', '/unpacked', '/differences', '/events', '/tasks'];
const TENANT_PATHS = ['/tenants'];

const HOME_URL = 'https://quest-on-demand.com/#/recovery/entraid/ng1/project-e71f6a66-3296-b0cf-e532-6fd2dea14c03/collection-ca9b03ea-578e-4277-b684-969fa2a34a9a/dashboard';

const navItems = [
  { key: 'home', icon: Home, label: 'Home', path: HOME_URL, external: true },
  { key: 'tenants', icon: Building2, label: 'Tenants', path: '/tenants' },
  { key: 'recover', icon: RotateCw, label: 'Recover', path: '/dashboard' },
  { key: 'settings', icon: Settings, label: 'Settings', path: '/setup' },
  { key: 'support', icon: HelpCircle, label: 'Support', path: '#' },
];

export default function Sidebar() {
  const location = useLocation();
  const navigate = useNavigate();

  const isRecoverActive = RECOVER_PATHS.some((p) => location.pathname === p || location.pathname.startsWith(p + '/'));
  const isTenantsActive = TENANT_PATHS.some((p) => location.pathname === p || location.pathname.startsWith(p + '/'));

  function isActive(item) {
    if (item.key === 'recover') return isRecoverActive;
    if (item.key === 'tenants') return isTenantsActive;
    if (item.key === 'settings') return location.pathname === '/setup';
    if (item.key === 'home') return location.pathname === '/';
    return false;
  }

  return (
    <div className="flex h-full flex-shrink-0">
      {/* Primary nav - 60px */}
      <div className="w-[60px] flex flex-col items-center pt-2" style={{ backgroundColor: '#1F1F1F' }}>
        {navItems.map((item) => {
          const active = isActive(item);
          const Icon = item.icon;
          return (
            <button
              key={item.key}
              onClick={() => {
              if (item.external) { window.location.href = item.path; }
              else if (item.path !== '#') { navigate(item.path); }
            }}
              className="w-full flex flex-col items-center py-3 relative cursor-pointer transition-colors"
              style={{ backgroundColor: active ? '#3C3C3C' : 'transparent' }}
            >
              {active && (
                <span
                  className="absolute left-0 top-0 bottom-0 w-[3px]"
                  style={{ backgroundColor: '#0096D6' }}
                />
              )}
              <Icon
                size={16}
                style={{ color: active ? '#0096D6' : '#999' }}
              />
              <span
                className="text-[10px] mt-1 font-medium"
                style={{ color: active ? '#fff' : '#888' }}
              >
                {item.label}
              </span>
            </button>
          );
        })}
      </div>

      {/* Secondary nav - 145px, only when recover is active */}
      {isRecoverActive && (
        <div className="w-[145px] flex flex-col pt-4" style={{ backgroundColor: '#3A3A3A' }}>
          <div className="relative px-3 py-2">
            <span
              className="absolute left-0 top-0 bottom-0 w-[3px]"
              style={{ backgroundColor: '#0096D6' }}
            />
            <span className="text-xs text-white font-medium pl-2">Microsoft Entra ID</span>
          </div>
        </div>
      )}
    </div>
  );
}
