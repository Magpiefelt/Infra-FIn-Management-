// Tests for project components
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ProjectCard, ProjectList, ProjectSummary } from '../components/projects/ProjectList.jsx'
import { renderWithProviders, mockProject } from './utils.jsx'

// Mock the hooks
vi.mock('../hooks/index.js', () => ({
  useProjects: vi.fn(() => ({
    projects: [mockProject],
    loading: false,
    error: null,
    refetch: vi.fn()
  }))
}))

describe('Project Components', () => {
  describe('ProjectCard', () => {
    it('should render project information correctly', () => {
      const onSelect = vi.fn()
      
      renderWithProviders(
        <ProjectCard project={mockProject} onSelect={onSelect} />
      )

      expect(screen.getByText('Test Project')).toBeInTheDocument()
      expect(screen.getByText('Test Contractor Ltd.')).toBeInTheDocument()
      expect(screen.getByText('Construction')).toBeInTheDocument()
      expect(screen.getByText('Test Region')).toBeInTheDocument()
      expect(screen.getByText('PM: Test Manager')).toBeInTheDocument()
    })

    it('should call onSelect when clicked', async () => {
      const user = userEvent.setup()
      const onSelect = vi.fn()
      
      renderWithProviders(
        <ProjectCard project={mockProject} onSelect={onSelect} />
      )

      const card = screen.getByText('Test Project').closest('.cursor-pointer')
      await user.click(card)

      expect(onSelect).toHaveBeenCalledWith()
    })

    it('should display budget information', () => {
      renderWithProviders(
        <ProjectCard project={mockProject} onSelect={vi.fn()} />
      )

      expect(screen.getByText('$1,000,000')).toBeInTheDocument()
      expect(screen.getByText(/Spent: \$500,000/)).toBeInTheDocument()
    })

    it('should display status badges', () => {
      renderWithProviders(
        <ProjectCard project={mockProject} onSelect={vi.fn()} />
      )

      expect(screen.getByText('Active')).toBeInTheDocument()
      expect(screen.getByText('Green')).toBeInTheDocument()
      expect(screen.getByText('Yellow')).toBeInTheDocument()
    })
  })

  describe('ProjectList', () => {
    it('should render list of projects', () => {
      renderWithProviders(
        <ProjectList onProjectSelect={vi.fn()} />
      )

      expect(screen.getByText('Test Project')).toBeInTheDocument()
    })

    it('should show loading state', () => {
      const { useProjects } = require('../hooks/index.js')
      useProjects.mockReturnValue({
        projects: [],
        loading: true,
        error: null,
        refetch: vi.fn()
      })

      renderWithProviders(
        <ProjectList onProjectSelect={vi.fn()} />
      )

      expect(screen.getByRole('status')).toBeInTheDocument()
    })

    it('should show error state', () => {
      const { useProjects } = require('../hooks/index.js')
      useProjects.mockReturnValue({
        projects: [],
        loading: false,
        error: 'Failed to load projects',
        refetch: vi.fn()
      })

      renderWithProviders(
        <ProjectList onProjectSelect={vi.fn()} />
      )

      expect(screen.getByText(/Failed to load projects/)).toBeInTheDocument()
    })

    it('should show empty state when no projects', () => {
      const { useProjects } = require('../hooks/index.js')
      useProjects.mockReturnValue({
        projects: [],
        loading: false,
        error: null,
        refetch: vi.fn()
      })

      renderWithProviders(
        <ProjectList onProjectSelect={vi.fn()} />
      )

      expect(screen.getByText('No projects found')).toBeInTheDocument()
    })
  })

  describe('ProjectSummary', () => {
    it('should render project summary cards', () => {
      renderWithProviders(
        <ProjectSummary project={mockProject} />
      )

      expect(screen.getByText('Total Budget')).toBeInTheDocument()
      expect(screen.getByText('TAF vs EAC')).toBeInTheDocument()
      expect(screen.getByText('Schedule Status')).toBeInTheDocument()
      expect(screen.getByText('Budget Status')).toBeInTheDocument()
    })

    it('should calculate and display variance', () => {
      renderWithProviders(
        <ProjectSummary project={mockProject} />
      )

      // EAC (1,050,000) - TAF (1,000,000) = +50,000 variance
      expect(screen.getByText(/\+\$50,000/)).toBeInTheDocument()
    })

    it('should not render when project is null', () => {
      const { container } = renderWithProviders(
        <ProjectSummary project={null} />
      )

      expect(container.firstChild).toBeNull()
    })
  })
})

