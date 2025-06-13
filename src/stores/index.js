// Centralized state management with Zustand
import { create } from 'zustand'
import { devtools } from 'zustand/middleware'
import { ProjectAPI, UserAPI } from '../services/apiService.js'

// Project store with pagination and API integration
export const useProjectStore = create(
  devtools(
    (set, get) => ({
      // State
      projects: [],
      selectedProject: null,
      loading: false,
      error: null,
      filter: 'all',
      
      // Pagination state
      currentPage: 1,
      pageSize: 10,
      totalProjects: 0,
      totalPages: 0,
      hasNext: false,
      hasPrev: false,

      // Actions
      setProjects: (projects) => set({ projects }),
      
      setSelectedProject: (project) => set({ selectedProject: project }),
      
      setLoading: (loading) => set({ loading }),
      
      setError: (error) => set({ error }),
      
      setFilter: (filter) => set({ filter }),

      // Pagination actions
      setCurrentPage: (page) => set({ currentPage: page }),
      
      setPageSize: (size) => {
        set({ pageSize: size, currentPage: 1 }) // Reset to first page when changing page size
        get().fetchProjects() // Refetch with new page size
      },

      setPaginationData: (pagination) => set({
        currentPage: pagination.page,
        pageSize: pagination.limit,
        totalProjects: pagination.total,
        totalPages: pagination.totalPages,
        hasNext: pagination.hasNext,
        hasPrev: pagination.hasPrev
      }),

      // Fetch projects from API with pagination
      fetchProjects: async (options = {}) => {
        const state = get()
        set({ loading: true, error: null })
        
        try {
          const apiOptions = {
            page: options.page || state.currentPage,
            limit: options.limit || state.pageSize,
            ...options
          }

          // Apply current filter
          if (state.filter !== 'all') {
            switch (state.filter) {
              case 'active':
                apiOptions.status = 'Active'
                break
              case 'completed':
                apiOptions.status = 'Completed'
                break
              case 'pending':
                apiOptions.reportStatus = 'Update Required'
                break
            }
          }

          const response = await ProjectAPI.getProjects(apiOptions)
          
          set({ 
            projects: response.data || [],
            loading: false
          })
          
          if (response.pagination) {
            get().setPaginationData(response.pagination)
          }
        } catch (error) {
          set({ 
            error: error.message,
            loading: false,
            projects: []
          })
        }
      },

      // Navigate to next page
      goToNextPage: () => {
        const state = get()
        if (state.hasNext) {
          const nextPage = state.currentPage + 1
          set({ currentPage: nextPage })
          get().fetchProjects({ page: nextPage })
        }
      },

      // Navigate to previous page
      goToPreviousPage: () => {
        const state = get()
        if (state.hasPrev) {
          const prevPage = state.currentPage - 1
          set({ currentPage: prevPage })
          get().fetchProjects({ page: prevPage })
        }
      },

      // Go to specific page
      goToPage: (page) => {
        const state = get()
        if (page >= 1 && page <= state.totalPages) {
          set({ currentPage: page })
          get().fetchProjects({ page })
        }
      },

      updateProject: async (id, updates) => {
        set({ loading: true, error: null })
        
        try {
          const response = await ProjectAPI.updateProject(id, updates)
          const updatedProject = response.data
          
          // Update projects list
          const projects = get().projects
          const updatedProjects = projects.map(project =>
            project.id === id ? updatedProject : project
          )
          set({ projects: updatedProjects })
          
          // Update selected project if it's the one being updated
          const selectedProject = get().selectedProject
          if (selectedProject && selectedProject.id === id) {
            set({ selectedProject: updatedProject })
          }
          
          set({ loading: false })
          return updatedProject
        } catch (error) {
          set({ error: error.message, loading: false })
          throw error
        }
      },

      addProject: async (projectData) => {
        set({ loading: true, error: null })
        
        try {
          const response = await ProjectAPI.createProject(projectData)
          const newProject = response.data
          
          // Refresh the projects list to maintain proper pagination
          await get().fetchProjects()
          
          set({ loading: false })
          return newProject
        } catch (error) {
          set({ error: error.message, loading: false })
          throw error
        }
      },

      removeProject: async (id) => {
        set({ loading: true, error: null })
        
        try {
          await ProjectAPI.deleteProject(id)
          
          // Refresh the projects list
          await get().fetchProjects()
          
          // Clear selected project if it's the one being removed
          const selectedProject = get().selectedProject
          if (selectedProject && selectedProject.id === id) {
            set({ selectedProject: null })
          }
          
          set({ loading: false })
        } catch (error) {
          set({ error: error.message, loading: false })
          throw error
        }
      },

      getProjectById: async (id) => {
        set({ loading: true, error: null })
        
        try {
          const response = await ProjectAPI.getProject(id)
          const project = response.data
          
          set({ selectedProject: project, loading: false })
          return project
        } catch (error) {
          set({ error: error.message, loading: false })
          return null
        }
      },

      // Upload Excel file for project
      uploadExcelFile: async (projectId, file) => {
        set({ loading: true, error: null })
        
        try {
          const response = await ProjectAPI.uploadExcel(projectId, file)
          const updatedProject = response.data
          
          // Update the project in the list
          const projects = get().projects
          const updatedProjects = projects.map(project =>
            project.id === projectId ? updatedProject : project
          )
          set({ projects: updatedProjects })
          
          // Update selected project if it's the one being updated
          const selectedProject = get().selectedProject
          if (selectedProject && selectedProject.id === projectId) {
            set({ selectedProject: updatedProject })
          }
          
          set({ loading: false })
          return response
        } catch (error) {
          set({ error: error.message, loading: false })
          throw error
        }
      },

      getFilteredProjects: () => {
        // This is now handled by the API, but keeping for compatibility
        return get().projects
      }
    }),
    {
      name: 'project-store'
    }
  )
)

// User/Auth store with API integration
export const useAuthStore = create(
  devtools(
    (set, get) => ({
      // State
      currentUser: {
        id: 5,
        name: "John Smith",
        role: "Vendor"
      },
      isAuthenticated: true,
      users: [],
      loadingUsers: false,

      // Actions
      setCurrentUser: (user) => set({ currentUser: user }),
      
      setRole: async (role) => {
        // Fetch users if not already loaded
        if (get().users.length === 0) {
          await get().fetchUsers()
        }
        
        const users = get().users
        const roleMap = {
          vendor: { role: "Vendor", defaultId: 5 },
          pm: { role: "Project Manager", defaultId: 1 },
          spm: { role: "Senior Project Manager", defaultId: 2 },
          director: { role: "Director", defaultId: 3 },
          admin: { role: "Director", defaultId: 4 } // Using Director role for admin
        }
        
        const targetRole = roleMap[role]?.role
        const defaultId = roleMap[role]?.defaultId
        
        // Find user with matching role or use default
        let user = users.find(u => u.role === targetRole)
        if (!user && defaultId) {
          user = users.find(u => u.id === defaultId)
        }
        
        if (user) {
          set({ currentUser: user })
        }
      },

      fetchUsers: async () => {
        set({ loadingUsers: true })
        try {
          const response = await UserAPI.getUsers()
          set({ users: response.data || [], loadingUsers: false })
        } catch (error) {
          console.error('Failed to fetch users:', error)
          set({ loadingUsers: false })
        }
      },

      login: (user) => set({ currentUser: user, isAuthenticated: true }),
      
      logout: () => set({ currentUser: null, isAuthenticated: false }),

      // Role-based project filtering
      getAccessibleProjects: (allProjects) => {
        const currentUser = get().currentUser
        if (!currentUser) return []

        // Directors and Senior Project Managers can see all projects
        if (['Director', 'Senior Project Manager'].includes(currentUser.role)) {
          return allProjects
        }

        // Project Managers can only see their own projects
        if (currentUser.role === 'Project Manager') {
          return allProjects.filter(project => project.ownerId === currentUser.id)
        }

        // Vendors have limited access - for demo purposes, show all projects
        if (currentUser.role === 'Vendor') {
          return allProjects // Changed from [] to allProjects for demo
        }

        return allProjects
      }
    }),
    {
      name: 'auth-store'
    }
  )
)

// UI state store
export const useUIStore = create(
  devtools(
    (set, get) => ({
      // State
      showPFMTExtractor: false,
      selectedProjectForPFMT: null,
      sidebarOpen: false,
      notifications: [],

      // Actions
      setShowPFMTExtractor: (show) => set({ showPFMTExtractor: show }),
      
      setSelectedProjectForPFMT: (project) => set({ selectedProjectForPFMT: project }),
      
      setSidebarOpen: (open) => set({ sidebarOpen: open }),
      
      addNotification: (notification) => {
        const notifications = get().notifications
        const newNotification = {
          id: Date.now(),
          timestamp: new Date().toISOString(),
          ...notification
        }
        set({ notifications: [...notifications, newNotification] })
      },
      
      removeNotification: (id) => {
        const notifications = get().notifications
        set({ notifications: notifications.filter(n => n.id !== id) })
      },
      
      clearNotifications: () => set({ notifications: [] })
    }),
    {
      name: 'ui-store'
    }
  )
)

