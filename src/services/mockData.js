// Enhanced mock data service with per-project storage and archiving
import { generateId } from '../utils/index.js'

// Project template for new projects
export const projectTemplate = {
  milestones: [
    {
      id: 1,
      name: "Project Initiation",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Project kickoff and team mobilization",
      notes: "",
      isNA: false,
      isRequired: true
    },
    {
      id: 2,
      name: "Design Completion",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Final design approval and documentation",
      notes: "",
      isNA: false,
      isRequired: true
    },
    {
      id: 3,
      name: "Construction Start",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Begin construction activities",
      notes: "",
      isNA: false,
      isRequired: true
    },
    {
      id: 4,
      name: "Structural Completion",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Complete structural work",
      notes: "",
      isNA: false,
      isRequired: true
    },
    {
      id: 5,
      name: "MEP Rough-in",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Mechanical, electrical, and plumbing rough-in",
      notes: "",
      isNA: false,
      isRequired: true
    },
    {
      id: 6,
      name: "Final Inspection",
      status: "pending",
      plannedDate: null,
      actualDate: null,
      baselineDate: null,
      description: "Final building inspection and certification",
      notes: "",
      isNA: false,
      isRequired: true
    }
  ],
  vendors: [
    {
      id: 1,
      name: "Primary Contractor",
      type: "General Contractor",
      contact: "",
      email: "",
      phone: "",
      status: "Active"
    }
  ],
  closeoutData: {
    finalInspectionDate: null,
    certificateOfOccupancy: null,
    warrantyPeriod: "12 months",
    deficiencyList: [],
    finalCost: null,
    lessonsLearned: "",
    projectRating: null
  },
  monthlyArchive: []
}

// Enhanced mock data with per-project storage
export const mockProjects = [
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
    // Per-project data
    milestones: [
      {
        id: 1,
        name: "Project Initiation",
        status: "completed",
        plannedDate: "2024-03-01",
        actualDate: "2024-02-28",
        baselineDate: "2024-03-01",
        description: "Project kickoff and team mobilization",
        notes: "Completed ahead of schedule",
        isNA: false,
        isRequired: true
      },
      {
        id: 2,
        name: "Design Completion",
        status: "completed",
        plannedDate: "2024-04-15",
        actualDate: "2024-04-12",
        baselineDate: "2024-04-15",
        description: "Final design approval and documentation",
        notes: "Minor revisions completed quickly",
        isNA: false,
        isRequired: true
      },
      {
        id: 3,
        name: "Construction Start",
        status: "completed",
        plannedDate: "2024-05-01",
        actualDate: "2024-05-03",
        baselineDate: "2024-05-01",
        description: "Begin construction activities",
        notes: "Slight delay due to permit processing",
        isNA: false,
        isRequired: true
      },
      {
        id: 4,
        name: "Structural Completion",
        status: "in-progress",
        plannedDate: "2024-07-15",
        actualDate: null,
        baselineDate: "2024-07-15",
        description: "Complete structural work",
        notes: "On track for completion",
        isNA: false,
        isRequired: true
      },
      {
        id: 5,
        name: "MEP Rough-in",
        status: "pending",
        plannedDate: "2024-08-30",
        actualDate: null,
        baselineDate: "2024-08-30",
        description: "Mechanical, electrical, and plumbing rough-in",
        notes: "",
        isNA: false,
        isRequired: true
      },
      {
        id: 6,
        name: "Final Inspection",
        status: "pending",
        plannedDate: "2024-11-15",
        actualDate: null,
        baselineDate: "2024-11-15",
        description: "Final building inspection and certification",
        notes: "",
        isNA: false,
        isRequired: true
      }
    ],
    vendors: [
      {
        id: 1,
        name: "ABC Construction Ltd.",
        type: "General Contractor",
        contact: "John Smith",
        email: "john@abcconstruction.com",
        phone: "(403) 555-0123",
        status: "Active"
      },
      {
        id: 2,
        name: "Steel Supply Co.",
        type: "Material Supplier",
        contact: "Jane Doe",
        email: "jane@steelsupply.com",
        phone: "(403) 555-0456",
        status: "Active"
      }
    ],
    closeoutData: {
      finalInspectionDate: null,
      certificateOfOccupancy: null,
      warrantyPeriod: "12 months",
      deficiencyList: [],
      finalCost: null,
      lessonsLearned: "",
      projectRating: null
    },
    monthlyArchive: []
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
    pfmtExtractedAt: "2024-05-15T10:30:00Z"
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
    seniorPmReviewed: false
  }
]

export const mockSubmissions = [
  {
    id: 1,
    purposeType: "Progress Payment",
    submissionDate: "2024-05-15",
    contractStartDate: "2024-03-15",
    expectedPaymentDate: "2024-06-01",
    status: "Approved",
    flagged: false,
    comments: "Foundation work completed as scheduled",
    constructionItems: ["Foundation", "Excavation", "Site Preparation"]
  },
  {
    id: 2,
    purposeType: "Change Order",
    submissionDate: "2024-05-20",
    contractStartDate: "2024-03-15",
    expectedPaymentDate: "2024-06-05",
    status: "Under Review",
    flagged: true,
    comments: "Additional electrical work required",
    constructionItems: ["Electrical", "HVAC Modifications"]
  }
]

export const mockMilestones = [
  {
    id: 1,
    name: "Project Initiation",
    status: "completed",
    plannedDate: "2024-03-01",
    actualDate: "2024-02-28",
    baselineDate: "2024-03-01",
    description: "Project kickoff and team mobilization",
    notes: "Completed ahead of schedule",
    isNA: false,
    isRequired: true
  },
  {
    id: 2,
    name: "Design Completion",
    status: "completed",
    plannedDate: "2024-04-15",
    actualDate: "2024-04-12",
    baselineDate: "2024-04-15",
    description: "Final design approval and documentation",
    notes: "Minor revisions completed quickly",
    isNA: false,
    isRequired: true
  },
  {
    id: 3,
    name: "Construction Start",
    status: "completed",
    plannedDate: "2024-05-01",
    actualDate: "2024-05-03",
    baselineDate: "2024-05-01",
    description: "Begin construction activities",
    notes: "Slight delay due to permit processing",
    isNA: false,
    isRequired: true
  },
  {
    id: 4,
    name: "Structural Completion",
    status: "in-progress",
    plannedDate: "2024-07-15",
    actualDate: null,
    baselineDate: "2024-07-15",
    description: "Complete structural work",
    notes: "On track for completion",
    isNA: false,
    isRequired: true
  },
  {
    id: 5,
    name: "MEP Rough-in",
    status: "pending",
    plannedDate: "2024-08-30",
    actualDate: null,
    baselineDate: "2024-08-30",
    description: "Mechanical, electrical, and plumbing rough-in",
    notes: "",
    isNA: false,
    isRequired: true
  },
  {
    id: 6,
    name: "Final Inspection",
    status: "pending",
    plannedDate: "2024-11-15",
    actualDate: null,
    baselineDate: "2024-11-15",
    description: "Final building inspection and certification",
    notes: "",
    isNA: false,
    isRequired: true
  }
]

// API simulation functions
export const getProjects = async (filter = 'all') => {
  // Simulate API delay
  await new Promise(resolve => setTimeout(resolve, 500))
  
  if (filter === 'all') {
    return mockProjects
  }
  
  return mockProjects.filter(project => {
    switch (filter) {
      case 'active':
        return project.status === 'Active'
      case 'completed':
        return project.status === 'Completed'
      case 'pending':
        return project.reportStatus === 'Update Required'
      default:
        return true
    }
  })
}

export const getProject = async (id) => {
  await new Promise(resolve => setTimeout(resolve, 300))
  return mockProjects.find(project => project.id === id)
}

export const updateProject = async (id, updates) => {
  await new Promise(resolve => setTimeout(resolve, 400))
  const projectIndex = mockProjects.findIndex(project => project.id === id)
  if (projectIndex !== -1) {
    mockProjects[projectIndex] = { ...mockProjects[projectIndex], ...updates }
    return mockProjects[projectIndex]
  }
  throw new Error('Project not found')
}

export const getSubmissions = async (projectId) => {
  await new Promise(resolve => setTimeout(resolve, 300))
  return mockSubmissions
}

export const getMilestones = async (projectId) => {
  await new Promise(resolve => setTimeout(resolve, 300))
  return mockMilestones
}

export const updateMilestone = async (milestoneId, updates) => {
  await new Promise(resolve => setTimeout(resolve, 200))
  const milestoneIndex = mockMilestones.findIndex(m => m.id === milestoneId)
  if (milestoneIndex !== -1) {
    mockMilestones[milestoneIndex] = { ...mockMilestones[milestoneIndex], ...updates }
    return mockMilestones[milestoneIndex]
  }
  throw new Error('Milestone not found')
}


// Post-approval archiving and workflow reset functions
export const archiveMonthlyData = async (projectId) => {
  await new Promise(resolve => setTimeout(resolve, 500))
  
  const projectIndex = mockProjects.findIndex(project => project.id === projectId)
  if (projectIndex === -1) {
    throw new Error('Project not found')
  }

  const project = mockProjects[projectIndex]
  const currentDate = new Date().toISOString()
  
  // Create archive entry for this month
  const archiveEntry = {
    id: generateId(),
    archivedDate: currentDate,
    reportingPeriod: new Date().toISOString().slice(0, 7), // YYYY-MM format
    data: {
      scheduleStatus: project.scheduleStatus,
      budgetStatus: project.budgetStatus,
      scheduleReasonCode: project.scheduleReasonCode,
      budgetReasonCode: project.budgetReasonCode,
      monthlyComments: project.monthlyComments,
      previousHighlights: project.previousHighlights,
      nextSteps: project.nextSteps,
      budgetVarianceExplanation: project.budgetVarianceExplanation,
      cashflowVarianceExplanation: project.cashflowVarianceExplanation,
      submittedBy: project.submittedBy,
      submittedDate: project.submittedDate,
      approvedBy: project.approvedBy,
      approvedDate: project.approvedDate,
      taf: project.taf,
      eac: project.eac,
      currentYearCashflow: project.currentYearCashflow,
      targetCashflow: project.targetCashflow,
      amountSpent: project.amountSpent
    }
  }

  // Add to monthly archive
  project.monthlyArchive.push(archiveEntry)

  // Reset workflow status for next cycle
  const updates = {
    reportStatus: "Update Required",
    scheduleReasonCode: "",
    budgetReasonCode: "",
    monthlyComments: "",
    previousHighlights: "",
    nextSteps: "",
    budgetVarianceExplanation: "",
    cashflowVarianceExplanation: "",
    submittedBy: null,
    submittedDate: null,
    approvedBy: null,
    approvedDate: null,
    directorApproved: false,
    seniorPmReviewed: false
  }

  // Apply updates
  Object.assign(project, updates)
  
  return {
    success: true,
    archiveEntry,
    message: `Monthly data archived successfully. Project status reset for next reporting cycle.`
  }
}

// Create new project function for PM workflow
export const createNewProject = async (projectData) => {
  await new Promise(resolve => setTimeout(resolve, 800))
  
  const newProject = {
    id: generateId(),
    name: projectData.name,
    description: projectData.description || "",
    preliminaryResourceRequirements: projectData.preliminaryResourceRequirements || "",
    programAssignment: projectData.programAssignment || "",
    clientMinistry: projectData.clientMinistry || "",
    projectType: projectData.projectType || "",
    createdBy: projectData.createdBy || "",
    createdDate: projectData.createdDate || new Date().toISOString(),
    contractor: "",
    startDate: null,
    submissions: 0,
    status: "Planning",
    reportStatus: "Not Started",
    phase: "Initiation",
    region: "",
    location: "",
    deliveryMethod: "Traditional",
    projectManager: projectData.projectManager || projectData.createdBy || "",
    seniorProjectManager: "",
    director: "",
    additionalTeam: [],
    totalBudget: 0,
    amountSpent: 0,
    taf: 0,
    eac: 0,
    currentYearCashflow: 0,
    targetCashflow: 0,
    lastPfmtUpdate: null,
    scheduleStatus: "Green",
    budgetStatus: "Green",
    scheduleReasonCode: "",
    budgetReasonCode: "",
    monthlyComments: "",
    previousHighlights: "",
    nextSteps: "Complete initial project planning and team assignment",
    budgetVarianceExplanation: "",
    cashflowVarianceExplanation: "",
    submittedBy: null,
    submittedDate: null,
    approvedBy: null,
    approvedDate: null,
    directorApproved: false,
    seniorPmReviewed: false,
    // Initialize with template data
    milestones: [
      {
        id: generateId(),
        name: "Project Initiation",
        status: "pending",
        plannedDate: null,
        actualDate: null,
        baselineDate: null,
        description: "Project kickoff and initial planning",
        notes: "",
        isNA: false,
        isRequired: true
      },
      {
        id: generateId(),
        name: "Scope Definition",
        status: "pending",
        plannedDate: null,
        actualDate: null,
        baselineDate: null,
        description: "Detailed scope and requirements definition",
        notes: "",
        isNA: false,
        isRequired: true
      },
      {
        id: generateId(),
        name: "Budget Approval",
        status: "pending",
        plannedDate: null,
        actualDate: null,
        baselineDate: null,
        description: "Budget finalization and approval",
        notes: "",
        isNA: false,
        isRequired: true
      }
    ],
    vendors: [
      {
        id: generateId(),
        name: "Primary Contractor",
        type: "General Contractor",
        contact: "",
        email: "",
        phone: "",
        status: "Pending Selection",
        contractValue: 0,
        notes: ""
      }
    ],
    closeoutData: {
      finalCost: 0,
      completionDate: "",
      lessonsLearned: "",
      clientSatisfaction: "",
      performanceMetrics: {}
    },
    monthlyArchive: []
  }

  // Add to projects array
  mockProjects.push(newProject)
  
  return newProject
}

// Generate project report
export const generateProjectReport = async (projectId, reportType = 'monthly') => {
  await new Promise(resolve => setTimeout(resolve, 1000))
  
  const project = mockProjects.find(p => p.id === projectId)
  if (!project) {
    throw new Error('Project not found')
  }

  const reportData = {
    id: generateId(),
    projectId: project.id,
    projectName: project.name,
    reportType,
    generatedDate: new Date().toISOString(),
    generatedBy: "System",
    data: {
      projectSummary: {
        name: project.name,
        contractor: project.contractor,
        phase: project.phase,
        region: project.region,
        projectManager: project.projectManager,
        startDate: project.startDate
      },
      financialSummary: {
        totalBudget: project.totalBudget,
        amountSpent: project.amountSpent,
        taf: project.taf,
        eac: project.eac,
        variance: project.eac - project.taf,
        currentYearCashflow: project.currentYearCashflow,
        targetCashflow: project.targetCashflow
      },
      statusSummary: {
        scheduleStatus: project.scheduleStatus,
        budgetStatus: project.budgetStatus,
        scheduleReasonCode: project.scheduleReasonCode,
        budgetReasonCode: project.budgetReasonCode,
        reportStatus: project.reportStatus
      },
      milestones: project.milestones,
      monthlyComments: project.monthlyComments,
      previousHighlights: project.previousHighlights,
      nextSteps: project.nextSteps,
      approvalHistory: {
        submittedBy: project.submittedBy,
        submittedDate: project.submittedDate,
        approvedBy: project.approvedBy,
        approvedDate: project.approvedDate,
        directorApproved: project.directorApproved
      },
      monthlyArchive: project.monthlyArchive
    }
  }

  return reportData
}

