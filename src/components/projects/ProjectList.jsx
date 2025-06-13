// Project-related components
import React, { useState } from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table.jsx'
import { Eye, Edit, FileText, Calendar, DollarSign, MapPin, Users } from 'lucide-react'
import { formatCurrency, formatDate, getStatusColor } from '../../utils/index.js'
import { useProjects } from '../../hooks/index.js'
import { LoadingSpinner, ErrorMessage, EmptyState } from '../shared/Layout.jsx'
import Pagination from '../Pagination.jsx'

export function ProjectList({ onProjectSelect, filter = 'all' }) {
  const { projects, loading, error, pagination, refetch } = useProjects(filter)

  if (loading) {
    return (
      <div className="flex justify-center py-8">
        <LoadingSpinner size="lg" />
      </div>
    )
  }

  if (error) {
    return (
      <ErrorMessage 
        message={`Failed to load projects: ${error}`}
        onRetry={refetch}
      />
    )
  }

  if (projects.length === 0 && pagination.totalProjects === 0) {
    return (
      <EmptyState
        title="No projects found"
        description="No projects match the current filter criteria."
        action={
          <Button onClick={refetch} variant="outline">
            Refresh
          </Button>
        }
      />
    )
  }

  return (
    <div className="space-y-6">
      {/* Projects list */}
      <div className="space-y-4">
        {projects.map((project) => (
          <ProjectCard 
            key={project.id} 
            project={project} 
            onSelect={() => onProjectSelect(project)}
          />
        ))}
      </div>

      {/* Pagination */}
      {pagination.totalPages > 1 && (
        <Pagination
          currentPage={pagination.currentPage}
          totalPages={pagination.totalPages}
          hasNext={pagination.hasNext}
          hasPrev={pagination.hasPrev}
          onNextPage={pagination.goToNextPage}
          onPrevPage={pagination.goToPreviousPage}
          onPageChange={pagination.goToPage}
          pageSize={pagination.pageSize}
          onPageSizeChange={pagination.setPageSize}
          totalItems={pagination.totalProjects}
          className="mt-6"
        />
      )}
    </div>
  )
}

export function ProjectCard({ project, onSelect }) {
  const budgetUtilization = (project.amountSpent / project.totalBudget) * 100

  return (
    <Card className="hover:shadow-md transition-shadow cursor-pointer" onClick={onSelect}>
      <CardHeader>
        <div className="flex items-start justify-between">
          <div className="flex-1">
            <CardTitle className="text-lg">{project.name}</CardTitle>
            <CardDescription className="mt-1">
              {project.contractor} • {project.phase}
            </CardDescription>
          </div>
          <div className="flex items-center space-x-2">
            <Badge variant="outline">{project.status}</Badge>
            <Badge 
              variant={project.reportStatus === 'Current' ? 'default' : 'destructive'}
            >
              {project.reportStatus}
            </Badge>
          </div>
        </div>
      </CardHeader>
      
      <CardContent>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="space-y-2">
            <div className="flex items-center text-sm text-gray-600">
              <MapPin className="h-4 w-4 mr-2" />
              {project.region}
            </div>
            <div className="flex items-center text-sm text-gray-600">
              <Calendar className="h-4 w-4 mr-2" />
              Started {formatDate(project.startDate)}
            </div>
            <div className="flex items-center text-sm text-gray-600">
              <Users className="h-4 w-4 mr-2" />
              PM: {project.projectManager}
            </div>
          </div>
          
          <div className="space-y-2">
            <div className="flex items-center text-sm">
              <DollarSign className="h-4 w-4 mr-2 text-gray-600" />
              <span className="font-medium">{formatCurrency(project.totalBudget)}</span>
            </div>
            <div className="text-sm text-gray-600">
              Spent: {formatCurrency(project.amountSpent)} ({budgetUtilization.toFixed(1)}%)
            </div>
            <div className="w-full bg-gray-200 rounded-full h-2">
              <div 
                className="bg-blue-600 h-2 rounded-full" 
                style={{ width: `${Math.min(budgetUtilization, 100)}%` }}
              ></div>
            </div>
          </div>
          
          <div className="space-y-2">
            <div className="flex items-center space-x-2">
              <span className="text-sm text-gray-600">Schedule:</span>
              <Badge className={getStatusColor(project.scheduleStatus)}>
                {project.scheduleStatus}
              </Badge>
            </div>
            <div className="flex items-center space-x-2">
              <span className="text-sm text-gray-600">Budget:</span>
              <Badge className={getStatusColor(project.budgetStatus)}>
                {project.budgetStatus}
              </Badge>
            </div>
            <Button variant="outline" size="sm" className="w-full">
              <Eye className="h-4 w-4 mr-2" />
              View Details
            </Button>
          </div>
        </div>
      </CardContent>
    </Card>
  )
}

export function ProjectTable({ onProjectSelect, filter = 'all' }) {
  const { projects, loading, error, pagination, refetch } = useProjects(filter)

  if (loading) {
    return (
      <div className="flex justify-center py-8">
        <LoadingSpinner size="lg" />
      </div>
    )
  }

  if (error) {
    return (
      <ErrorMessage 
        message={`Failed to load projects: ${error}`}
        onRetry={refetch}
      />
    )
  }

  if (projects.length === 0 && pagination.totalProjects === 0) {
    return (
      <EmptyState
        title="No projects to display"
        description="There are no projects available to show in the table."
        action={
          <Button onClick={refetch} variant="outline">
            Refresh
          </Button>
        }
      />
    )
  }

  return (
    <div className="space-y-6">
      {/* Projects table */}
      <div className="border rounded-lg overflow-hidden">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Project Name</TableHead>
              <TableHead>Contractor</TableHead>
              <TableHead>Phase</TableHead>
              <TableHead>Budget</TableHead>
              <TableHead>Schedule</TableHead>
              <TableHead>Budget Status</TableHead>
              <TableHead>Actions</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {projects.map((project) => (
              <TableRow key={project.id} className="hover:bg-gray-50">
                <TableCell>
                  <div>
                    <div className="font-medium">{project.name}</div>
                    <div className="text-sm text-gray-600">{project.region}</div>
                  </div>
                </TableCell>
                <TableCell>{project.contractor}</TableCell>
                <TableCell>
                  <Badge variant="outline">{project.phase}</Badge>
                </TableCell>
                <TableCell>
                  <div>
                    <div className="font-medium">{formatCurrency(project.totalBudget)}</div>
                    <div className="text-sm text-gray-600">
                      Spent: {formatCurrency(project.amountSpent)}
                    </div>
                  </div>
                </TableCell>
                <TableCell>
                  <Badge className={getStatusColor(project.scheduleStatus)}>
                    {project.scheduleStatus}
                  </Badge>
                </TableCell>
                <TableCell>
                  <Badge className={getStatusColor(project.budgetStatus)}>
                    {project.budgetStatus}
                  </Badge>
                </TableCell>
                <TableCell>
                  <Button 
                    variant="outline" 
                    size="sm"
                    onClick={() => onProjectSelect(project)}
                  >
                    <Eye className="h-4 w-4 mr-2" />
                    View
                  </Button>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {/* Pagination */}
      {pagination.totalPages > 1 && (
        <Pagination
          currentPage={pagination.currentPage}
          totalPages={pagination.totalPages}
          hasNext={pagination.hasNext}
          hasPrev={pagination.hasPrev}
          onNextPage={pagination.goToNextPage}
          onPrevPage={pagination.goToPreviousPage}
          onPageChange={pagination.goToPage}
          pageSize={pagination.pageSize}
          onPageSizeChange={pagination.setPageSize}
          totalItems={pagination.totalProjects}
          className="mt-6"
        />
      )}
    </div>
  )
}

export function ProjectSummary({ project }) {
  if (!project) return null

  const budgetUtilization = (project.amountSpent / project.totalBudget) * 100
  const variance = project.eac - project.taf

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-medium text-gray-600">
            Total Budget
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-2xl font-bold">{formatCurrency(project.totalBudget)}</div>
          <div className="text-sm text-gray-600">
            Spent: {formatCurrency(project.amountSpent)} ({budgetUtilization.toFixed(1)}%)
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-medium text-gray-600">
            TAF vs EAC
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-2xl font-bold">{formatCurrency(project.taf)}</div>
          <div className={`text-sm ${variance > 0 ? 'text-red-600' : 'text-green-600'}`}>
            EAC: {formatCurrency(project.eac)} ({variance > 0 ? '+' : ''}{formatCurrency(variance)})
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-medium text-gray-600">
            Schedule Status
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Badge className={getStatusColor(project.scheduleStatus)}>
            {project.scheduleStatus}
          </Badge>
          {project.scheduleReasonCode && (
            <div className="text-sm text-gray-600 mt-2">
              {project.scheduleReasonCode}
            </div>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-medium text-gray-600">
            Budget Status
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Badge className={getStatusColor(project.budgetStatus)}>
            {project.budgetStatus}
          </Badge>
          {project.budgetReasonCode && (
            <div className="text-sm text-gray-600 mt-2">
              {project.budgetReasonCode}
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  )
}

