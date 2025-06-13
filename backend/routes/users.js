// User routes
import express from 'express'
import { UserController } from '../controllers/userController.js'

const router = express.Router()

// Routes
router.get('/', UserController.getAllUsers)
router.get('/:id', UserController.getUserById)
router.get('/role/:role', UserController.getUsersByRole)

export default router

