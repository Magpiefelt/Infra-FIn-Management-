// Test utilities and helpers
import React from 'react'
import { render } from '@testing-library/react'
import { BrowserRouter } from 'react-router-dom'

// Custom render function that includes providers
export function renderWithProviders(ui, options = {}) {
  const { initialEntries = ['/'], ...renderOptions } = options

  function Wrapper({ children }) {
    return (
      <BrowserRouter>
        {children}
      </BrowserRouter>
    )
  }

  return render(ui, { wrapper: Wrapper, ...renderOptions })
}

// Mock data for testing
export const mockProject = {
  id: "test-project-1",
  name: "Test Project",
  contractor: "Test Contractor Ltd.",
  startDate: "2024-01-01",
  submissions: 5,
  status: "Active",
  reportStatus: "Current",
  description: "Test project description",
  phase: "Construction",
  region: "Test Region",
  location: "123 Test Street",
  deliveryMethod: "Design-Build",
  projectManager: "Test Manager",
  seniorProjectManager: "Test Senior Manager",
  director: "Test Director",
  additionalTeam: ["Test Engineer"],
  totalBudget: 1000000,
  amountSpent: 500000,
  taf: 1000000,
  eac: 1050000,
  currentYearCashflow: 400000,
  targetCashflow: 450000,
  lastPfmtUpdate: null,
  scheduleStatus: "Green",
  budgetStatus: "Yellow",
  scheduleReasonCode: "",
  budgetReasonCode: "Test reason",
  monthlyComments: "Test comments",
  previousHighlights: "Test highlights",
  nextSteps: "Test next steps",
  budgetVarianceExplanation: "Test explanation",
  cashflowVarianceExplanation: "Test explanation",
  submittedBy: null,
  submittedDate: null,
  approvedBy: null,
  approvedDate: null,
  directorApproved: false,
  seniorPmReviewed: false
}

export const mockUser = {
  name: "Test User",
  role: "pm"
}

// Helper functions for testing
export const waitForLoadingToFinish = () => {
  return new Promise(resolve => setTimeout(resolve, 0))
}

export const createMockStore = (initialState = {}) => {
  return {
    projects: [mockProject],
    selectedProject: null,
    loading: false,
    error: null,
    ...initialState
  }
}

