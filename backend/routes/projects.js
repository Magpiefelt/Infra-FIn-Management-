// Project routes
import express from 'express'
import multer from 'multer'
import path from 'path'
import { fileURLToPath } from 'url'
import { ProjectController } from '../controllers/projectController.js'

const router = express.Router()

// Get __dirname equivalent for ES modules
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, '..', 'uploads')
    cb(null, uploadDir)
  },
  filename: (req, file, cb) => {
    // Generate unique filename with timestamp
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)
    const ext = path.extname(file.originalname)
    cb(null, `pfmt-${uniqueSuffix}${ext}`)
  }
})

// File filter for Excel files only
const fileFilter = (req, file, cb) => {
  const allowedTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel.sheet.macroEnabled.12' // .xlsm
  ]
  
  const allowedExtensions = ['.xlsx', '.xlsm']
  const fileExtension = path.extname(file.originalname).toLowerCase()
  
  if (allowedTypes.includes(file.mimetype) || allowedExtensions.includes(fileExtension)) {
    cb(null, true)
  } else {
    cb(new Error('Only Excel files (.xlsx, .xlsm) are allowed'), false)
  }
}

// Configure multer with options
const upload = multer({
  storage: storage,
  fileFilter: fileFilter,
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB limit
  }
})

// Create uploads directory if it doesn't exist
import fs from 'fs'
const uploadsDir = path.join(__dirname, '..', 'uploads')
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true })
}

// Routes
router.get('/', ProjectController.getAllProjects)
router.get('/:id', ProjectController.getProjectById)
router.post('/', ProjectController.createProject)
router.put('/:id', ProjectController.updateProject)
router.delete('/:id', ProjectController.deleteProject)

// Excel upload route
router.post('/:id/excel', upload.single('file'), ProjectController.uploadExcel)

// PFMT Excel upload route with data extraction
router.post('/:id/pfmt-excel', upload.single('file'), ProjectController.uploadPFMTExcel)

export default router

