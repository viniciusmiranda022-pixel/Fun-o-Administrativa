import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import Dashboard from './pages/Dashboard.jsx';
import Backups from './pages/Backups.jsx';
import UnpackedObjects from './pages/UnpackedObjects.jsx';
import Differences from './pages/Differences.jsx';
import Events from './pages/Events.jsx';
import Tasks from './pages/Tasks.jsx';

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
      </Route>
    </Routes>
  );
}
