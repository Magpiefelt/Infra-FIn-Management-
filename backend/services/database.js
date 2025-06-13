// Database service using lowdb for JSON-based persistence
import { Low } from 'lowdb'
import { JSONFileSync } from 'lowdb/node'
import { v4 as uuidv4 } from 'uuid'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Database file path
const dbFile = path.join(__dirname, '..', 'db.json')

// Database adapter
const adapter = new JSONFileSync(dbFile)
const db = new Low(adapter, {})

// Default data structure
const defaultData = {
  projects: [
    {
      id: "88:36:D3",
      name: "Calgary Elementary School Renovation",
      contractor: "ABC Construction Ltd.",
      startDate: "2024-03-15",
      submissions: 12,
      status: "Active",
      reportStatus: "Update Required",
      description: "Complete renovation of elementary school facilities including classroom upgrades and safety improvements",
      phase: "Construction",
      region: "Calgary",
      location: "123 School Street, Calgary, AB T2P 1A1",
      deliveryMethod: "Design-Build",
      projectManager: "Sarah Johnson",
      seniorProjectManager: "Michael Brown",
      director: "Lisa Wilson",
      additionalTeam: ["John Doe - Engineer", "Jane Smith - Architect"],
      totalBudget: 2450000,
      amountSpent: 1680000,
      taf: 2450000,
      eac: 2520000,
      currentYearCashflow: 1200000,
      targetCashflow: 1150000,
      lastPfmtUpdate: null,
      scheduleStatus: "Green",
      budgetStatus: "Yellow",
      scheduleReasonCode: "",
      budgetReasonCode: "Material Supply Delays",
      monthlyComments: "Foundation work completed ahead of schedule but material delays affecting budget",
      previousHighlights: "Completed foundation work 2 weeks early",
      nextSteps: "Proceed with framing and electrical rough-in",
      budgetVarianceExplanation: "Steel prices increased by 15% due to supply chain issues",
      cashflowVarianceExplanation: "Additional payments required for expedited material delivery",
      submittedBy: null,
      submittedDate: null,
      approvedBy: null,
      approvedDate: null,
      directorApproved: false,
      seniorPmReviewed: false,
      ownerId: 1,
      createdAt: "2024-03-01T08:00:00.000Z",
      updatedAt: "2024-06-12T10:30:00.000Z"
    },
    {
      id: "77:24:B1",
      name: "Edmonton Hospital Wing Addition",
      contractor: "MediConstruct Inc.",
      startDate: "2024-01-10",
      submissions: 8,
      status: "Active",
      reportStatus: "Current",
      description: "New wing addition to Edmonton General Hospital with specialized medical equipment installation",
      phase: "Design Development",
      region: "Edmonton",
      location: "456 Health Avenue, Edmonton, AB T5K 2M3",
      deliveryMethod: "Traditional",
      projectManager: "Michael Brown",
      seniorProjectManager: "Lisa Wilson",
      director: "David Chen",
      additionalTeam: ["Dr. Sarah Lee - Medical Consultant", "Tom Wilson - MEP Engineer"],
      totalBudget: 5200000,
      amountSpent: 1560000,
      taf: 5200000,
      eac: 5100000,
      currentYearCashflow: 2100000,
      targetCashflow: 2200000,
      lastPfmtUpdate: "2024-05-15T10:30:00Z",
      scheduleStatus: "Green",
      budgetStatus: "Green",
      scheduleReasonCode: "",
      budgetReasonCode: "",
      monthlyComments: "Project progressing well with all milestones on track",
      previousHighlights: "Completed design phase approval",
      nextSteps: "Begin construction phase and equipment procurement",
      budgetVarianceExplanation: "",
      cashflowVarianceExplanation: "",
      submittedBy: "Michael Brown",
      submittedDate: "2024-05-20T14:00:00Z",
      approvedBy: "Lisa Wilson",
      approvedDate: "2024-05-21T09:15:00Z",
      directorApproved: true,
      seniorPmReviewed: true,
      pfmtFileName: "Edmonton-Hospital-PFMT-May2024.xlsx",
      pfmtExtractedAt: "2024-05-15T10:30:00Z",
      ownerId: 2,
      createdAt: "2024-01-01T08:00:00.000Z",
      updatedAt: "2024-05-21T09:15:00.000Z"
    },
    {
      id: "99:48:C7",
      name: "Red Deer Community Center",
      contractor: "Community Builders Ltd.",
      startDate: "2024-02-20",
      submissions: 15,
      status: "Active",
      reportStatus: "Review Required",
      description: "Multi-purpose community center with gymnasium, meeting rooms, and recreational facilities",
      phase: "Construction",
      region: "Red Deer",
      location: "789 Community Drive, Red Deer, AB T4N 1A5",
      deliveryMethod: "Design-Build",
      projectManager: "Sarah Johnson",
      seniorProjectManager: "Michael Brown",
      director: "Lisa Wilson",
      additionalTeam: ["Mark Thompson - Structural Engineer", "Lisa Chen - Interior Designer"],
      totalBudget: 3800000,
      amountSpent: 2280000,
      taf: 3800000,
      eac: 3950000,
      currentYearCashflow: 1900000,
      targetCashflow: 1800000,
      lastPfmtUpdate: null,
      scheduleStatus: "Yellow",
      budgetStatus: "Red",
      scheduleReasonCode: "Weather Delays",
      budgetReasonCode: "Change Orders",
      monthlyComments: "Weather delays in March affected schedule. Additional change orders for accessibility improvements",
      previousHighlights: "Structural work completed",
      nextSteps: "Interior finishing and equipment installation",
      budgetVarianceExplanation: "Client requested additional accessibility features adding $150K to budget",
      cashflowVarianceExplanation: "Accelerated payments to maintain schedule",
      submittedBy: null,
      submittedDate: null,
      approvedBy: null,
      approvedDate: null,
      directorApproved: false,
      seniorPmReviewed: false,
      ownerId: 1,
      createdAt: "2024-02-15T08:00:00.000Z",
      updatedAt: "2024-06-10T14:20:00.000Z"
    }
  ],
  users: [
    { id: 1, name: "Sarah Johnson", role: "Project Manager", email: "sarah.johnson@company.com" },
    { id: 2, name: "Michael Brown", role: "Senior Project Manager", email: "michael.brown@company.com" },
    { id: 3, name: "Lisa Wilson", role: "Director", email: "lisa.wilson@company.com" },
    { id: 4, name: "David Chen", role: "Director", email: "david.chen@company.com" },
    { id: 5, name: "John Smith", role: "Vendor", email: "john.smith@vendor.com" }
  ]
}

// Initialize database
export function initializeDatabase() {
  try {
    db.read()
    
    // Initialize with default data if database is empty
    if (!db.data || Object.keys(db.data).length === 0) {
      db.data = defaultData
      db.write()
      console.log('Database initialized with default data')
    }
    
    return db
  } catch (error) {
    console.error('Failed to initialize database:', error)
    throw error
  }
}

// Get database instance
export function getDatabase() {
  return db
}

// Helper functions for common operations
export function getAllProjects(filters = {}) {
  db.read()
  let projects = db.data.projects || []
  
  // Apply filters
  if (filters.ownerId) {
    projects = projects.filter(p => p.ownerId === parseInt(filters.ownerId))
  }
  
  if (filters.status) {
    projects = projects.filter(p => p.status === filters.status)
  }
  
  if (filters.reportStatus) {
    projects = projects.filter(p => p.reportStatus === filters.reportStatus)
  }
  
  // Sort by createdAt descending (newest first)
  projects.sort((a, b) => {
    const dateA = a.createdAt ? new Date(a.createdAt).getTime() : 0
    const dateB = b.createdAt ? new Date(b.createdAt).getTime() : 0
    return dateB - dateA
  })
  
  return projects
}

export function getProjectById(id) {
  db.read()
  return db.data.projects?.find(p => p.id === id) || null
}

export function createProject(projectData) {
  db.read()
  
  const newProject = {
    id: uuidv4(),
    ...projectData,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  }
  
  db.data.projects = db.data.projects || []
  db.data.projects.push(newProject)
  db.write()
  
  return newProject
}

export function updateProject(id, updates) {
  db.read()
  
  const projectIndex = db.data.projects?.findIndex(p => p.id === id)
  if (projectIndex === -1) {
    return null
  }
  
  db.data.projects[projectIndex] = {
    ...db.data.projects[projectIndex],
    ...updates,
    updatedAt: new Date().toISOString()
  }
  
  db.write()
  return db.data.projects[projectIndex]
}

export function deleteProject(id) {
  db.read()
  
  const projectIndex = db.data.projects?.findIndex(p => p.id === id)
  if (projectIndex === -1) {
    return false
  }
  
  db.data.projects.splice(projectIndex, 1)
  db.write()
  return true
}

export function getAllUsers() {
  db.read()
  return db.data.users || []
}

export function getUserById(id) {
  db.read()
  return db.data.users?.find(u => u.id === parseInt(id)) || null
}

