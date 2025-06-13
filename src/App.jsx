// Simplified refactored App.jsx using modular components and centralized state
import React, { useState } from 'react'
import { BrowserRouter as Router, Routes, Route, Navigate, useNavigate } from 'react-router-dom'
import { PFMTDataExtractor } from './components/PFMTDataExtractor.jsx'
import { Header, Layout } from './components/shared/Layout.jsx'
import { ErrorBoundary } from './components/shared/ErrorBoundary.jsx'
import { ProjectList } from './components/projects/ProjectList.jsx'
import { ProjectManagerDashboard } from './components/projects/AddNewProject.jsx'
import { useAuth, usePFMTExtractor } from './hooks/index.js'
import './App.css'

// Home page with navigation tiles
function HomePage() {
  const { currentUser } = useAuth()
  const navigate = useNavigate()
  
  const navigationTiles = [
    { title: 'My Projects', description: 'View and manage your assigned projects', path: '/projects' },
    { title: 'All Projects', description: 'Browse all projects in the system', path: '/projects' },
    { title: 'Gate Meetings', description: 'Schedule and manage gate meetings', path: '/meetings' },
    { title: 'PFMT Library', description: 'Access PFMT templates and resources', path: '/library' },
    { title: 'Reporting', description: 'Generate and view project reports', path: '/reports' },
    { title: 'Settings', description: 'Manage your account and preferences', path: '/settings' }
  ]

  const handleTileClick = (path) => {
    navigate(path)
  }

  return (
    <div className="p-6">
      <div className="mb-8">
        <h1 className="text-3xl font-bold text-gray-900 mb-2">
          Welcome back, {currentUser.name}
        </h1>
        <p className="text-gray-600">
          {currentUser.role === 'vendor' ? 
            'Manage your project submissions and track progress.' :
            'Access your project management tools and resources.'
          }
        </p>
      </div>
      
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {navigationTiles.map((tile) => (
          <div 
            key={tile.title} 
            className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 hover:shadow-md transition-shadow cursor-pointer"
            onClick={() => handleTileClick(tile.path)}
          >
            <h3 className="text-lg font-semibold text-gray-900 mb-2">{tile.title}</h3>
            <p className="text-gray-600 text-sm">{tile.description}</p>
          </div>
        ))}
      </div>

      {/* Project Manager Dashboard - only show for PMs and above */}
      {['pm', 'spm', 'director', 'admin'].includes(currentUser.role) && (
        <div className="mt-8">
          <ProjectManagerDashboard projects={[]} onProjectCreated={() => {}} />
        </div>
      )}
    </div>
  )
}

// Vendor portal component
function VendorPortal() {
  return (
    <div className="p-6">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Vendor Submission Portal</h2>
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <p className="text-gray-600">
          Submit your project information and track submission status.
        </p>
      </div>
    </div>
  )
}

// Staff portal component
function StaffPortal() {
  const { currentUser } = useAuth()
  
  if (currentUser.role === 'vendor') {
    return <Navigate to="/vendor" replace />
  }
  
  return <HomePage />
}

// Projects page component
function ProjectsPage() {
  const [selectedProject, setSelectedProject] = React.useState(null)
  
  if (selectedProject) {
    return (
      <div className="p-6">
        <button 
          onClick={() => setSelectedProject(null)}
          className="mb-4 text-blue-600 hover:text-blue-800"
        >
          ← Back to Projects
        </button>
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h2 className="text-2xl font-bold text-gray-900 mb-4">{selectedProject.name}</h2>
          <p className="text-gray-600">{selectedProject.description}</p>
        </div>
      </div>
    )
  }
  
  return (
    <div className="p-6">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Projects</h2>
      <ProjectList onProjectSelect={setSelectedProject} />
    </div>
  )
}

// PFMT Extractor Page
function PFMTExtractorPage() {
  const [selectedProject, setSelectedProject] = useState(null)
  const [extractedData, setExtractedData] = useState(null)

  const handleDataExtracted = (data) => {
    setExtractedData(data)
    // Here you could integrate with the project creation workflow
    console.log('Extracted PFMT data:', data)
  }

  return (
    <div className="p-6">
      <div className="mb-6">
        <h1 className="text-3xl font-bold text-gray-900 mb-2">PFMT Data Extractor</h1>
        <p className="text-gray-600">
          Upload and extract data from PFMT Excel files to create new projects or update existing ones.
        </p>
      </div>
      
      <div className="max-w-4xl">
        <PFMTDataExtractor
          project={selectedProject}
          onDataExtracted={handleDataExtracted}
          onClose={() => setSelectedProject(null)}
        />
      </div>
      
      {extractedData && (
        <div className="mt-8">
          <h2 className="text-xl font-semibold text-gray-900 mb-4">Extracted Data Preview</h2>
          <div className="bg-gray-50 p-4 rounded-lg">
            <pre className="text-sm overflow-auto">
              {JSON.stringify(extractedData, null, 2)}
            </pre>
          </div>
        </div>
      )}
    </div>
  )
}

// Main app content component
function AppContent() {
  const navigate = useNavigate()
  const { currentUser, changeRole } = useAuth()
  const { 
    showPFMTExtractor, 
    selectedProjectForPFMT, 
    closePFMTExtractor, 
    handlePFMTDataExtracted 
  } = usePFMTExtractor()

  const handleRoleChange = (newRole) => {
    changeRole(newRole)
    
    // Navigate based on role
    if (newRole === 'vendor') {
      navigate('/vendor', { replace: true })
    } else {
      navigate('/staff', { replace: true })
    }
  }

  return (
    <Layout>
      <Header onRoleChange={handleRoleChange} />
      <main>
        <Routes>
          <Route 
            path="/" 
            element={
              currentUser.role === 'vendor' ? 
              <Navigate to="/vendor" replace /> : 
              <Navigate to="/staff" replace />
            } 
          />
          <Route path="/vendor" element={<VendorPortal />} />
          <Route path="/staff" element={<StaffPortal />} />
          <Route path="/projects" element={<ProjectsPage />} />
          <Route path="/projects/all" element={<ProjectsPage />} />
          <Route path="/meetings" element={<div className="p-6">Gate Meetings - Coming Soon</div>} />
          <Route path="/library" element={<PFMTExtractorPage />} />
          <Route path="/extractor" element={<PFMTExtractorPage />} />
          <Route path="/reports" element={<div className="p-6">Reporting - Coming Soon</div>} />
          <Route path="/settings" element={<div className="p-6">Settings - Coming Soon</div>} />
        </Routes>
      </main>
      
      {/* PFMT Data Extractor Dialog */}
      {showPFMTExtractor && selectedProjectForPFMT && (
        <PFMTDataExtractor
          project={selectedProjectForPFMT}
          onDataExtracted={handlePFMTDataExtracted}
          onClose={closePFMTExtractor}
        />
      )}
    </Layout>
  )
}

// Main App component with error boundary
function App() {
  return (
    <ErrorBoundary fallbackMessage="The PFMT application encountered an unexpected error. Please refresh the page to continue.">
      <Router>
        <AppContent />
      </Router>
    </ErrorBoundary>
  )
}

export default App

