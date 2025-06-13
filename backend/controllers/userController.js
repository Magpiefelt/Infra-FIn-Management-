// User controller for handling user-related HTTP requests
import { UserService } from '../services/userService.js'

export class UserController {
  // GET /api/users
  static async getAllUsers(req, res) {
    try {
      const users = UserService.getAllUsers()
      
      res.json({
        success: true,
        data: users
      })
    } catch (error) {
      console.error('Error getting users:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to retrieve users',
        message: error.message
      })
    }
  }
  
  // GET /api/users/:id
  static async getUserById(req, res) {
    try {
      const { id } = req.params
      const user = UserService.getUserById(parseInt(id))
      
      if (!user) {
        return res.status(404).json({
          success: false,
          error: 'User not found'
        })
      }
      
      res.json({
        success: true,
        data: user
      })
    } catch (error) {
      console.error('Error getting user:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to retrieve user',
        message: error.message
      })
    }
  }
  
  // GET /api/users/role/:role
  static async getUsersByRole(req, res) {
    try {
      const { role } = req.params
      const users = UserService.getUsersByRole(role)
      
      res.json({
        success: true,
        data: users
      })
    } catch (error) {
      console.error('Error getting users by role:', error)
      res.status(500).json({
        success: false,
        error: 'Failed to retrieve users by role',
        message: error.message
      })
    }
  }
}

