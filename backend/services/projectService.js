// Project service layer for business logic
import * as db from './database.js'
import * as XLSX from 'xlsx'
import fs from 'fs/promises'

export class ProjectService {
  // Get all projects with pagination and filtering
  static getProjects(options = {}) {
    const {
      page = 1,
      limit = 10,
      ownerId,
      status,
      reportStatus
    } = options
    
    // Get filtered projects
    const allProjects = db.getAllProjects({
      ownerId,
      status,
      reportStatus
    })
    
    // Calculate pagination
    const total = allProjects.length
    const totalPages = Math.ceil(total / limit)
    const startIndex = (page - 1) * limit
    const endIndex = startIndex + limit
    
    // Get page data
    const projects = allProjects.slice(startIndex, endIndex)
    
    return {
      projects,
      pagination: {
        page: parseInt(page),
        limit: parseInt(limit),
        total,
        totalPages,
        hasNext: page < totalPages,
        hasPrev: page > 1
      }
    }
  }
  
  // Get single project by ID
  static getProjectById(id) {
    return db.getProjectById(id)
  }
  
  // Create new project
  static createProject(projectData) {
    // Validate required fields
    const requiredFields = ['name', 'description']
    for (const field of requiredFields) {
      if (!projectData[field]) {
        throw new Error(`Missing required field: ${field}`)
      }
    }
    
    // Set default values
    const defaultProject = {
      status: 'Active',
      reportStatus: 'Update Required',
      phase: 'Planning',
      submissions: 0,
      totalBudget: 0,
      amountSpent: 0,
      taf: 0,
      eac: 0,
      currentYearCashflow: 0,
      targetCashflow: 0,
      scheduleStatus: 'Green',
      budgetStatus: 'Green',
      scheduleReasonCode: '',
      budgetReasonCode: '',
      monthlyComments: '',
      previousHighlights: '',
      nextSteps: '',
      budgetVarianceExplanation: '',
      cashflowVarianceExplanation: '',
      submittedBy: null,
      submittedDate: null,
      approvedBy: null,
      approvedDate: null,
      directorApproved: false,
      seniorPmReviewed: false,
      lastPfmtUpdate: null,
      pfmtFileName: null,
      pfmtExtractedAt: null,
      additionalTeam: []
    }
    
    const newProject = {
      ...defaultProject,
      ...projectData
    }
    
    return db.createProject(newProject)
  }
  
  // Update project
  static updateProject(id, updates) {
    const project = db.getProjectById(id)
    if (!project) {
      throw new Error('Project not found')
    }
    
    return db.updateProject(id, updates)
  }
  
  // Delete project
  static deleteProject(id) {
    const project = db.getProjectById(id)
    if (!project) {
      throw new Error('Project not found')
    }
    
    return db.deleteProject(id)
  }
  
  // Process Excel file upload
  static async processExcelUpload(projectId, filePath, fileName) {
    try {
      // Verify project exists
      const project = db.getProjectById(projectId)
      if (!project) {
        throw new Error('Project not found')
      }
      
      // Read and parse Excel file
      const workbook = XLSX.readFile(filePath)
      
      // Look for the "SP Fields" sheet or use the first sheet
      let sheetName = 'SP Fields'
      if (!workbook.SheetNames.includes(sheetName)) {
        sheetName = workbook.SheetNames[0]
        if (!sheetName) {
          throw new Error('No worksheets found in Excel file')
        }
      }
      
      const worksheet = workbook.Sheets[sheetName]
      
      // Helper function to safely get cell value
      const getCellValue = (cellAddress) => {
        const cell = worksheet[cellAddress]
        if (!cell) return null
        
        // Handle different cell types
        if (cell.t === 'n') return cell.v // number
        if (cell.t === 's') return cell.v // string
        if (cell.t === 'b') return cell.v // boolean
        if (cell.t === 'd') return cell.v // date
        
        return cell.v
      }
      
      // Extract financial data from specific cells
      // These cell addresses should match the PFMT template structure
      const extractedData = {
        taf: getCellValue('C5') || 0, // Total Approved Funding
        eac: getCellValue('C6') || 0, // Estimate at Completion
        currentYearCashflow: getCellValue('C7') || 0, // Current Year Cashflow
        currentYearTarget: getCellValue('C8') || 0, // Current Year Target
        // Add more fields as needed based on PFMT template
      }
      
      // Validate extracted data
      const numericFields = ['taf', 'eac', 'currentYearCashflow', 'currentYearTarget']
      for (const field of numericFields) {
        if (extractedData[field] && typeof extractedData[field] !== 'number') {
          // Try to convert to number
          const numValue = parseFloat(extractedData[field])
          if (!isNaN(numValue)) {
            extractedData[field] = numValue
          } else {
            extractedData[field] = 0
          }
        }
      }
      
      // Calculate variances
      const tafEacVariance = extractedData.eac - extractedData.taf
      const cashflowVariance = extractedData.currentYearCashflow - extractedData.currentYearTarget
      
      // Prepare update data
      const updateData = {
        taf: extractedData.taf,
        eac: extractedData.eac,
        currentYearCashflow: extractedData.currentYearCashflow,
        targetCashflow: extractedData.currentYearTarget, // Use consistent field name
        lastPfmtUpdate: new Date().toISOString(),
        pfmtFileName: fileName,
        pfmtExtractedAt: new Date().toISOString(),
        reportStatus: 'Current', // Update status to indicate current data
        // Store calculated variances if needed
        tafEacVariance,
        cashflowVariance
      }
      
      // Update project with extracted data
      const updatedProject = db.updateProject(projectId, updateData)
      
      // Clean up uploaded file
      try {
        await fs.unlink(filePath)
      } catch (cleanupError) {
        console.warn('Failed to clean up uploaded file:', cleanupError)
      }
      
      return {
        project: updatedProject,
        extractedData: {
          ...extractedData,
          tafEacVariance,
          cashflowVariance,
          fileName,
          extractedAt: updateData.pfmtExtractedAt
        }
      }
      
    } catch (error) {
      // Clean up uploaded file on error
      try {
        await fs.unlink(filePath)
      } catch (cleanupError) {
        console.warn('Failed to clean up uploaded file after error:', cleanupError)
      }
      
      throw error
    }
  }
}

