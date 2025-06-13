// Project controller for handling HTTP requests
import { ProjectService } from '../services/projectService.js'
import { UserService } from '../services/userService.js'

export class ProjectController {
  // GET /api/projects
  static async getAllProjects(req, res) {
    try {
      const {
        page = 1,
        limit = 10,
        ownerId,
        status,
        reportStatus
      } = req.query
      
      const options = {
        page: parseInt(page),
        limit: parseInt(limit)
      }
      
      // Add filters if provided
      if (ownerId) options.ownerId = ownerId
      if (status) options.status = status
      if (reportStatus) options.reportStatus = reportStatus
      
      const result = ProjectService.getProjects(options)
      
      res.json({
        success: true,
        data: result.projects,
        pagination: result.pagination
      })
    } catch (error) {
      console.error('Error getting projects:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to retrieve projects',
        message: error.message
      })
    }
  }
  
  // GET /api/projects/:id
  static async getProjectById(req, res) {
    try {
      const { id } = req.params
      const project = ProjectService.getProjectById(id)
      
      if (!project) {
        return res.status(404).json({
          success: false,
          error: 'Project not found'
        })
      }
      
      res.json({
        success: true,
        data: project
      })
    } catch (error) {
      console.error('Error getting project:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to retrieve project',
        message: error.message
      })
    }
  }
  
  // POST /api/projects
  static async createProject(req, res) {
    try {
      const projectData = req.body
      
      // Validate required fields
      if (!projectData.name || !projectData.description) {
        return res.status(400).json({
          success: false,
          error: 'Missing required fields',
          message: 'Name and description are required'
        })
      }
      
      const newProject = ProjectService.createProject(projectData)
      
      res.status(201).json({
        success: true,
        data: newProject,
        message: 'Project created successfully'
      })
    } catch (error) {
      console.error('Error creating project:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to create project',
        message: error.message
      })
    }
  }
  
  // PUT /api/projects/:id
  static async updateProject(req, res) {
    try {
      const { id } = req.params
      const updates = req.body
      
      const updatedProject = ProjectService.updateProject(id, updates)
      
      res.json({
        success: true,
        data: updatedProject,
        message: 'Project updated successfully'
      })
    } catch (error) {
      console.error('Error updating project:', error)
      
      if (error.message === 'Project not found') {
        return res.status(404).json({
          success: false,
          error: 'Project not found'
        })
      }
      
      res.status(500).json({
        success: false,
        error: 'Failed to update project',
        message: error.message
      })
    }
  }
  
  // DELETE /api/projects/:id
  static async deleteProject(req, res) {
    try {
      const { id } = req.params
      
      ProjectService.deleteProject(id)
      
      res.json({
        success: true,
        message: 'Project deleted successfully'
      })
    } catch (error) {
      console.error('Error deleting project:', error)
      
      if (error.message === 'Project not found') {
        return res.status(404).json({
          success: false,
          error: 'Project not found'
        })
      }
      
      res.status(500).json({
        success: false,
        error: 'Failed to delete project',
        message: error.message
      })
    }
  }
  
  // POST /api/projects/:id/excel
  static async uploadExcel(req, res) {
    try {
      const { id } = req.params
      
      if (!req.file) {
        return res.status(400).json({
          success: false,
          error: 'No file uploaded',
          message: 'Please select an Excel file to upload'
        })
      }
      
      const result = await ProjectService.processExcelUpload(
        id,
        req.file.path,
        req.file.originalname
      )
      
      res.json({
        success: true,
        data: result.project,
        extractedData: result.extractedData,
        message: 'Excel file processed successfully'
      })
    } catch (error) {
      console.error('Error processing Excel upload:', error)
      
      if (error.message === 'Project not found') {
        return res.status(404).json({
          success: false,
          error: 'Project not found'
        })
      }
      
      if (error.message.includes('No worksheets found')) {
        return res.status(400).json({
          success: false,
          error: 'Invalid Excel file',
          message: 'The uploaded file does not contain any worksheets'
        })
      }
      
      res.status(500).json({
        success: false,
        error: 'Failed to process Excel file',
        message: error.message
      })
    }
  }
}

