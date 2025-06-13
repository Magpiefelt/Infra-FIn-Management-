// Tests for Zustand stores
import { describe, it, expect, beforeEach } from 'vitest'
import { act, renderHook } from '@testing-library/react'
import { useProjectStore, useAuthStore, useUIStore } from '../stores/index.js'

describe('Zustand Stores', () => {
  describe('useProjectStore', () => {
    beforeEach(() => {
      // Reset store state before each test
      useProjectStore.setState({
        projects: [],
        selectedProject: null,
        loading: false,
        error: null,
        filter: 'all'
      })
    })

    it('should initialize with default state', () => {
      const { result } = renderHook(() => useProjectStore())
      
      expect(result.current.projects).toEqual([])
      expect(result.current.selectedProject).toBeNull()
      expect(result.current.loading).toBe(false)
      expect(result.current.error).toBeNull()
      expect(result.current.filter).toBe('all')
    })

    it('should update projects', () => {
      const { result } = renderHook(() => useProjectStore())
      const testProjects = [{ id: '1', name: 'Test Project' }]

      act(() => {
        result.current.setProjects(testProjects)
      })

      expect(result.current.projects).toEqual(testProjects)
    })

    it('should update individual project', () => {
      const { result } = renderHook(() => useProjectStore())
      const testProjects = [
        { id: '1', name: 'Test Project', status: 'Active' },
        { id: '2', name: 'Another Project', status: 'Active' }
      ]

      act(() => {
        result.current.setProjects(testProjects)
      })

      act(() => {
        result.current.updateProject('1', { status: 'Completed' })
      })

      expect(result.current.projects[0].status).toBe('Completed')
      expect(result.current.projects[1].status).toBe('Active')
    })

    it('should add new project', () => {
      const { result } = renderHook(() => useProjectStore())
      const newProject = { id: '1', name: 'New Project' }

      act(() => {
        result.current.addProject(newProject)
      })

      expect(result.current.projects).toContain(newProject)
    })

    it('should remove project', () => {
      const { result } = renderHook(() => useProjectStore())
      const testProjects = [
        { id: '1', name: 'Test Project' },
        { id: '2', name: 'Another Project' }
      ]

      act(() => {
        result.current.setProjects(testProjects)
      })

      act(() => {
        result.current.removeProject('1')
      })

      expect(result.current.projects).toHaveLength(1)
      expect(result.current.projects[0].id).toBe('2')
    })

    it('should get project by id', () => {
      const { result } = renderHook(() => useProjectStore())
      const testProjects = [
        { id: '1', name: 'Test Project' },
        { id: '2', name: 'Another Project' }
      ]

      act(() => {
        result.current.setProjects(testProjects)
      })

      const project = result.current.getProjectById('1')
      expect(project).toEqual(testProjects[0])
    })
  })

  describe('useAuthStore', () => {
    beforeEach(() => {
      useAuthStore.setState({
        currentUser: { name: "John Smith", role: "vendor" },
        isAuthenticated: true
      })
    })

    it('should initialize with default user', () => {
      const { result } = renderHook(() => useAuthStore())
      
      expect(result.current.currentUser.name).toBe("John Smith")
      expect(result.current.currentUser.role).toBe("vendor")
      expect(result.current.isAuthenticated).toBe(true)
    })

    it('should change user role', () => {
      const { result } = renderHook(() => useAuthStore())

      act(() => {
        result.current.setRole('pm')
      })

      expect(result.current.currentUser.role).toBe('pm')
      expect(result.current.currentUser.name).toBe('Sarah Johnson')
    })

    it('should handle login', () => {
      const { result } = renderHook(() => useAuthStore())
      const newUser = { name: 'Test User', role: 'admin' }

      act(() => {
        result.current.login(newUser)
      })

      expect(result.current.currentUser).toEqual(newUser)
      expect(result.current.isAuthenticated).toBe(true)
    })

    it('should handle logout', () => {
      const { result } = renderHook(() => useAuthStore())

      act(() => {
        result.current.logout()
      })

      expect(result.current.currentUser).toBeNull()
      expect(result.current.isAuthenticated).toBe(false)
    })
  })

  describe('useUIStore', () => {
    beforeEach(() => {
      useUIStore.setState({
        showPFMTExtractor: false,
        selectedProjectForPFMT: null,
        sidebarOpen: false,
        notifications: []
      })
    })

    it('should initialize with default UI state', () => {
      const { result } = renderHook(() => useUIStore())
      
      expect(result.current.showPFMTExtractor).toBe(false)
      expect(result.current.selectedProjectForPFMT).toBeNull()
      expect(result.current.sidebarOpen).toBe(false)
      expect(result.current.notifications).toEqual([])
    })

    it('should toggle PFMT extractor', () => {
      const { result } = renderHook(() => useUIStore())

      act(() => {
        result.current.setShowPFMTExtractor(true)
      })

      expect(result.current.showPFMTExtractor).toBe(true)
    })

    it('should add notification', () => {
      const { result } = renderHook(() => useUIStore())
      const notification = { type: 'success', message: 'Test message' }

      act(() => {
        result.current.addNotification(notification)
      })

      expect(result.current.notifications).toHaveLength(1)
      expect(result.current.notifications[0]).toMatchObject(notification)
      expect(result.current.notifications[0]).toHaveProperty('id')
      expect(result.current.notifications[0]).toHaveProperty('timestamp')
    })

    it('should remove notification', () => {
      const { result } = renderHook(() => useUIStore())
      
      act(() => {
        result.current.addNotification({ type: 'info', message: 'Test' })
      })

      const notificationId = result.current.notifications[0].id

      act(() => {
        result.current.removeNotification(notificationId)
      })

      expect(result.current.notifications).toHaveLength(0)
    })

    it('should clear all notifications', () => {
      const { result } = renderHook(() => useUIStore())
      
      act(() => {
        result.current.addNotification({ type: 'info', message: 'Test 1' })
        result.current.addNotification({ type: 'info', message: 'Test 2' })
      })

      expect(result.current.notifications).toHaveLength(2)

      act(() => {
        result.current.clearNotifications()
      })

      expect(result.current.notifications).toHaveLength(0)
    })
  })
})

