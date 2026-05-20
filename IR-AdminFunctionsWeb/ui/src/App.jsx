import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import Dashboard from './pages/Dashboard.jsx';
import Backups from './pages/Backups.jsx';
import UnpackedObjects from './pages/UnpackedObjects.jsx';
import Differences from './pages/Differences.jsx';
import Events from './pages/Events.jsx';
import Tasks from './pages/Tasks.jsx';
import Tenants from './pages/Tenants.jsx';
import TenantConsents from './pages/TenantConsents.jsx';
import Setup from './pages/Setup.jsx';
import ManageBackups from './pages/ManageBackups.jsx';
import ManageRestores from './pages/ManageRestores.jsx';
import UnpackBackup from './pages/UnpackBackup.jsx';

export default function App() {
  return (
    <Routes>
      <Route element={<Layout />}>
        <Route path="/" element={<Navigate to="/dashboard" replace />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/backups" element={<Backups />} />
        <Route path="/unpacked" element={<UnpackedObjects />} />
        <Route path="/differences" element={<Differences />} />
        <Route path="/events" element={<Events />} />
        <Route path="/tasks" element={<Tasks />} />
        <Route path="/tenants" element={<Tenants />} />
        <Route path="/tenants/:tenantId/consents" element={<TenantConsents />} />
        <Route path="/setup" element={<Setup />} />
        <Route path="/manage-backups" element={<ManageBackups />} />
        <Route path="/manage-restores" element={<ManageRestores />} />
        <Route path="/unpack" element={<UnpackBackup />} />
      </Route>
    </Routes>
  );
}
