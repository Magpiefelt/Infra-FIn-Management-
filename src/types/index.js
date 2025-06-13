// Shared types and interfaces
export interface User {
  name: string
  role: 'vendor' | 'pm' | 'spm' | 'director' | 'admin'
}

export interface Project {
  id: string
  name: string
  contractor: string
  startDate: string
  submissions: number
  status: string
  reportStatus: string
  description: string
  phase: string
  region: string
  location: string
  deliveryMethod: string
  projectManager: string
  seniorProjectManager: string
  director: string
  additionalTeam: string[]
  totalBudget: number
  amountSpent: number
  taf: number
  eac: number
  currentYearCashflow: number
  targetCashflow: number
  lastPfmtUpdate: string | null
  scheduleStatus: 'Green' | 'Yellow' | 'Red'
  budgetStatus: 'Green' | 'Yellow' | 'Red'
  scheduleReasonCode: string
  budgetReasonCode: string
  monthlyComments: string
  previousHighlights: string
  nextSteps: string
  budgetVarianceExplanation: string
  cashflowVarianceExplanation: string
  submittedBy: string | null
  submittedDate: string | null
  approvedBy: string | null
  approvedDate: string | null
  directorApproved: boolean
  seniorPmReviewed: boolean
  pfmtFileName?: string
  pfmtExtractedAt?: string
}

export interface Submission {
  id: number
  purposeType: string
  submissionDate: string
  contractStartDate: string
  expectedPaymentDate: string
  status: string
  flagged: boolean
  comments: string
  constructionItems: string[]
}

export interface Milestone {
  id: number
  name: string
  status: 'completed' | 'in-progress' | 'pending' | 'overdue' | 'na'
  plannedDate: string
  actualDate: string | null
  baselineDate: string
  description: string
  notes: string
  isNA: boolean
  isRequired: boolean
}

export interface PFMTData {
  taf: number | null
  eac: number | null
  currentYearCashflow: number | null
  currentYearTarget: number | null
  projectId?: string | null
  lastUpdated: string
  fileName: string
  fileSize: number
  extractedAt: string
  availableSheets: string[]
  tafEacVariance?: number
  cashflowVariance?: number
}

export interface ValidationResult {
  isValid: boolean
  issues: string[]
  warnings: string[]
  hasWarnings: boolean
}

export type StatusColor = 'Green' | 'Yellow' | 'Red'
export type UserRole = 'vendor' | 'pm' | 'spm' | 'director' | 'admin'

