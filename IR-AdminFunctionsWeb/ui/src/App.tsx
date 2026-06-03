import { lazy, Suspense } from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import { AppProvider } from './context/AppContext.jsx';
import Layout from './components/layout/Layout.jsx';
import ErrorBoundary from './components/common/ErrorBoundary.jsx';

// Route-level code splitting: each page is loaded only when first visited
const Dashboard     = lazy(() => import('./pages/Dashboard.jsx'));
const Backups       = lazy(() => import('./pages/Backups.jsx'));
const UnpackedObjects = lazy(() => import('./pages/UnpackedObjects.jsx'));
const Differences   = lazy(() => import('./pages/Differences.jsx'));
const Events        = lazy(() => import('./pages/Events.jsx'));
const Tasks         = lazy(() => import('./pages/Tasks.jsx'));
const Tenants       = lazy(() => import('./pages/Tenants.jsx'));
const TenantConsents = lazy(() => import('./pages/TenantConsents.jsx'));
const Setup         = lazy(() => import('./pages/Setup.jsx'));

function PageLoader() {
  return (
    <div className="flex items-center justify-center h-full">
      <div className="w-6 h-6 border-2 border-[#0078A8] border-t-transparent rounded-full animate-spin" />
    </div>
  );
}

export default function App() {
  return (
    <AppProvider>
      <ErrorBoundary>
        <Routes>
          <Route element={<Layout />}>
            <Route path="/" element={<Navigate to="/dashboard" replace />} />
            <Route path="/dashboard"        element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Dashboard /></ErrorBoundary></Suspense>} />
            <Route path="/backups"          element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Backups /></ErrorBoundary></Suspense>} />
            <Route path="/unpacked"         element={<Suspense fallback={<PageLoader />}><ErrorBoundary><UnpackedObjects /></ErrorBoundary></Suspense>} />
            <Route path="/differences"      element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Differences /></ErrorBoundary></Suspense>} />
            <Route path="/events"           element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Events /></ErrorBoundary></Suspense>} />
            <Route path="/tasks"            element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Tasks /></ErrorBoundary></Suspense>} />
            <Route path="/tenants"          element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Tenants /></ErrorBoundary></Suspense>} />
            <Route path="/tenants/:tenantId/consents" element={<Suspense fallback={<PageLoader />}><ErrorBoundary><TenantConsents /></ErrorBoundary></Suspense>} />
            <Route path="/setup"            element={<Suspense fallback={<PageLoader />}><ErrorBoundary><Setup /></ErrorBoundary></Suspense>} />
          </Route>
        </Routes>
      </ErrorBoundary>
    </AppProvider>
  );
}
