// User service layer for user management
import * as db from './database.js'

export class UserService {
  // Get all users
  static getAllUsers() {
    return db.getAllUsers()
  }
  
  // Get user by ID
  static getUserById(id) {
    return db.getUserById(id)
  }
  
  // Get users by role
  static getUsersByRole(role) {
    const users = db.getAllUsers()
    return users.filter(user => user.role === role)
  }
  
  // Validate user role for authorization
  static validateRole(userRole, requiredRoles) {
    if (!Array.isArray(requiredRoles)) {
      requiredRoles = [requiredRoles]
    }
    
    return requiredRoles.includes(userRole)
  }
  
  // Get role hierarchy level (for permission checking)
  static getRoleLevel(role) {
    const roleLevels = {
      'Vendor': 1,
      'Project Manager': 2,
      'Senior Project Manager': 3,
      'Director': 4,
      'Admin': 5
    }
    
    return roleLevels[role] || 0
  }
  
  // Check if user can access project (based on role and ownership)
  static canAccessProject(user, project) {
    // Directors and Admins can access all projects
    if (['Director', 'Admin'].includes(user.role)) {
      return true
    }
    
    // Senior Project Managers can access all projects
    if (user.role === 'Senior Project Manager') {
      return true
    }
    
    // Project Managers can only access their own projects
    if (user.role === 'Project Manager') {
      return project.ownerId === user.id
    }
    
    // Vendors have limited access (would need additional logic)
    if (user.role === 'Vendor') {
      // For now, vendors can't access projects through this API
      return false
    }
    
    return false
  }
  
  // Get projects accessible to user
  static getAccessibleProjects(user, filters = {}) {
    const allProjects = db.getAllProjects(filters)
    
    // Filter projects based on user role
    return allProjects.filter(project => this.canAccessProject(user, project))
  }
}

