// Custom hooks for business logic
import { useState, useEffect, useCallback } from 'react'
import { useProjectStore, useAuthStore, useUIStore } from '../stores/index.js'

// Hook for project management with pagination
export const useProjects = (filter = 'all') => {
  try {
    const store = useProjectStore()
    const authStore = useAuthStore()

    // Destructure with fallbacks to avoid undefined errors
    const {
      projects = [],
      loading = false,
      error = null,
      currentPage = 1,
      pageSize = 10,
      totalProjects = 0,
      totalPages = 0,
      hasNext = false,
      hasPrev = false,
      setFilter,
      fetchProjects,
      goToNextPage,
      goToPreviousPage,
      goToPage,
      setPageSize,
      updateProject,
      addProject,
      removeProject
    } = store || {}

    const { getAccessibleProjects } = authStore || {}

    // Set filter and fetch projects when filter changes
    useEffect(() => {
      if (setFilter && fetchProjects) {
        setFilter(filter)
        fetchProjects()
      }
    }, [filter, setFilter, fetchProjects])

    // Fetch projects on initial mount
    useEffect(() => {
      console.log('useEffect for fetchProjects triggered, fetchProjects available:', !!fetchProjects)
      if (fetchProjects) {
        console.log('Calling fetchProjects from useEffect')
        fetchProjects()
      }
    }, [fetchProjects])

    // Apply role-based filtering to projects
    const accessibleProjects = getAccessibleProjects ? getAccessibleProjects(projects) : projects

    const pagination = {
      currentPage,
      pageSize,
      totalProjects,
      totalPages,
      hasNext,
      hasPrev,
      goToNextPage,
      goToPreviousPage,
      goToPage,
      setPageSize
    }

    return {
      projects: accessibleProjects,
      loading,
      error,
      pagination,
      refetch: fetchProjects,
      updateProject,
      addProject,
      removeProject
    }
  } catch (error) {
    console.error('Error in useProjects hook:', error)
    // Return fallback values if store is not available
    return {
      projects: [],
      loading: false,
      error: null,
      pagination: {
        currentPage: 1,
        pageSize: 10,
        totalProjects: 0,
        totalPages: 0,
        hasNext: false,
        hasPrev: false,
        goToNextPage: () => {},
        goToPreviousPage: () => {},
        goToPage: () => {},
        setPageSize: () => {}
      },
      refetch: () => {},
      updateProject: () => {},
      addProject: () => {},
      removeProject: () => {}
    }
  }
}

// Hook for individual project management
export const useProject = (projectId) => {
  // const { getProjectById, updateProject } = useProjectStore()
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  // Temporary simplified implementation
  const project = null

  const updateProjectData = useCallback(async (updates) => {
    setLoading(true)
    setError(null)
    try {
      // await updateProjectAPI(projectId, updates)
      // updateProject(projectId, updates)
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }, [projectId])

  return {
    project,
    loading,
    error,
    updateProject: updateProjectData,
    refetch: () => {}
  }
}

// Hook for authentication
export const useAuth = () => {
  try {
    const authStore = useAuthStore()
    const { currentUser, isAuthenticated, setRole, login, logout, fetchUsers } = authStore || {}

    const handleRoleChange = useCallback(async (newRole) => {
      if (setRole) {
        await setRole(newRole)
      }
    }, [setRole])

    // Fetch users on first use
    useEffect(() => {
      if (fetchUsers) {
        fetchUsers()
      }
    }, [fetchUsers])

    return {
      currentUser: currentUser || {
        id: 1,
        name: 'Demo User',
        role: 'pm',
        email: 'demo@example.com'
      },
      isAuthenticated: isAuthenticated !== undefined ? isAuthenticated : true,
      changeRole: handleRoleChange,
      login: login || (() => {}),
      logout: logout || (() => {})
    }
  } catch (error) {
    console.error('Error in useAuth hook:', error)
    // Return fallback values if store is not available
    return {
      currentUser: {
        id: 1,
        name: 'Demo User',
        role: 'pm',
        email: 'demo@example.com'
      },
      isAuthenticated: true,
      changeRole: () => {},
      login: () => {},
      logout: () => {}
    }
  }
}

// Hook for PFMT data extraction
export const usePFMTExtractor = () => {
  try {
    const uiStore = useUIStore()
    const projectStore = useProjectStore()
    
    const {
      showPFMTExtractor,
      selectedProjectForPFMT,
      setShowPFMTExtractor,
      setSelectedProjectForPFMT
    } = uiStore || {}
    
    const { uploadExcelFile } = projectStore || {}

    const openPFMTExtractor = useCallback((project) => {
      if (setSelectedProjectForPFMT && setShowPFMTExtractor) {
        setSelectedProjectForPFMT(project)
        setShowPFMTExtractor(true)
      }
    }, [setSelectedProjectForPFMT, setShowPFMTExtractor])

    const closePFMTExtractor = useCallback(() => {
      if (setShowPFMTExtractor && setSelectedProjectForPFMT) {
        setShowPFMTExtractor(false)
        setSelectedProjectForPFMT(null)
      }
    }, [setShowPFMTExtractor, setSelectedProjectForPFMT])

    const handlePFMTDataExtracted = useCallback(async (file) => {
      if (selectedProjectForPFMT && file && uploadExcelFile) {
        try {
          const result = await uploadExcelFile(selectedProjectForPFMT.id, file)
          return result
        } catch (error) {
          console.error('Failed to upload Excel file:', error)
          throw error
        }
      }
    }, [selectedProjectForPFMT, uploadExcelFile])

    return {
      showPFMTExtractor: showPFMTExtractor || false,
      selectedProjectForPFMT: selectedProjectForPFMT || null,
      openPFMTExtractor,
      closePFMTExtractor,
      handlePFMTDataExtracted
    }
  } catch (error) {
    console.error('Error in usePFMTExtractor hook:', error)
    // Return fallback values if store is not available
    return {
      showPFMTExtractor: false,
      selectedProjectForPFMT: null,
      openPFMTExtractor: () => {},
      closePFMTExtractor: () => {},
      handlePFMTDataExtracted: () => {}
    }
  }
}

// Hook for notifications
export const useNotifications = () => {
  return {
    notifications: [],
    showSuccess: () => {},
    showError: () => {},
    showWarning: () => {},
    showInfo: () => {},
    removeNotification: () => {},
    clearNotifications: () => {}
  }
}

// Hook for form validation
export const useFormValidation = (initialValues, validationRules) => {
  const [values, setValues] = useState(initialValues)
  const [errors, setErrors] = useState({})
  const [touched, setTouched] = useState({})

  const validateField = useCallback((name, value) => {
    const rules = validationRules[name]
    if (!rules) return null

    for (const rule of rules) {
      const error = rule(value)
      if (error) return error
    }
    return null
  }, [validationRules])

  const handleChange = useCallback((name, value) => {
    setValues(prev => ({ ...prev, [name]: value }))
    
    if (touched[name]) {
      const error = validateField(name, value)
      setErrors(prev => ({ ...prev, [name]: error }))
    }
  }, [touched, validateField])

  const handleBlur = useCallback((name) => {
    setTouched(prev => ({ ...prev, [name]: true }))
    const error = validateField(name, values[name])
    setErrors(prev => ({ ...prev, [name]: error }))
  }, [values, validateField])

  const validateAll = useCallback(() => {
    const newErrors = {}
    let isValid = true

    Object.keys(validationRules).forEach(name => {
      const error = validateField(name, values[name])
      if (error) {
        newErrors[name] = error
        isValid = false
      }
    })

    setErrors(newErrors)
    setTouched(Object.keys(validationRules).reduce((acc, key) => {
      acc[key] = true
      return acc
    }, {}))

    return isValid
  }, [values, validationRules, validateField])

  const reset = useCallback(() => {
    setValues(initialValues)
    setErrors({})
    setTouched({})
  }, [initialValues])

  return {
    values,
    errors,
    touched,
    handleChange,
    handleBlur,
    validateAll,
    reset,
    isValid: Object.keys(errors).length === 0
  }
}
