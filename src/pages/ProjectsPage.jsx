// Enhanced ProjectsPage with comprehensive project detail view
import React, { useState, useEffect } from 'react'
import { useSearchParams } from 'react-router-dom'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs.jsx'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog.jsx'
import { Input } from '@/components/ui/input.jsx'
import { Label } from '@/components/ui/label.jsx'
import { Textarea } from '@/components/ui/textarea.jsx'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select.jsx'
import { ArrowLeft, Edit, Upload, FileText, Users, Building, CheckCircle, Calendar, DollarSign, MapPin, Filter } from 'lucide-react'
import { ProjectList, ProjectSummary } from '../components/projects/ProjectList.jsx'
import { ProjectMilestones, ProjectVendors } from '../components/projects/ProjectDataManagement.jsx'
import { PFMTDataExtractor } from '../components/PFMTDataExtractor.jsx'
import { useProjects, useNotifications, usePFMTExtractor } from '../hooks/index.js'
import { formatCurrency, formatDate, getStatusColor } from '../utils/index.js'

// Project Team Management Component
function ProjectTeam({ project, onProjectUpdate }) {
  const [showAddDialog, setShowAddDialog] = useState(false)
  const [newTeamMember, setNewTeamMember] = useState({
    name: '',
    role: '',
    email: '',
    phone: '',
    department: ''
  })
  const { showSuccess, showError } = useNotifications()

  const handleAddTeamMember = async () => {
    try {
      const teamMember = {
        id: Date.now(),
        ...newTeamMember
      }
      
      const updatedTeam = [...(project.team || []), teamMember]
      await onProjectUpdate(project.id, { team: updatedTeam })
      
      setNewTeamMember({ name: '', role: '', email: '', phone: '', department: '' })
      setShowAddDialog(false)
      showSuccess('Team member added successfully')
    } catch (error) {
      showError('Failed to add team member: ' + error.message)
    }
  }

  const handleRemoveTeamMember = async (memberId) => {
    try {
      const updatedTeam = project.team.filter(m => m.id !== memberId)
      await onProjectUpdate(project.id, { team: updatedTeam })
      showSuccess('Team member removed successfully')
    } catch (error) {
      showError('Failed to remove team member: ' + error.message)
    }
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div>
            <CardTitle className="flex items-center space-x-2">
              <Users className="h-5 w-5" />
              <span>Project Team</span>
            </CardTitle>
            <CardDescription>Manage project team members and contacts</CardDescription>
          </div>
          <Dialog open={showAddDialog} onOpenChange={setShowAddDialog}>
            <DialogTrigger asChild>
              <Button size="sm">
                <Users className="h-4 w-4 mr-2" />
                Add Team Member
              </Button>
            </DialogTrigger>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Add Team Member</DialogTitle>
                <DialogDescription>
                  Add a new team member to this project
                </DialogDescription>
              </DialogHeader>
              <div className="space-y-4">
                <div>
                  <Label htmlFor="member-name">Name</Label>
                  <Input
                    id="member-name"
                    value={newTeamMember.name}
                    onChange={(e) => setNewTeamMember(prev => ({ ...prev, name: e.target.value }))}
                    placeholder="Enter team member name"
                  />
                </div>
                <div>
                  <Label htmlFor="member-role">Role</Label>
                  <Select value={newTeamMember.role} onValueChange={(value) => setNewTeamMember(prev => ({ ...prev, role: value }))}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select role" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Project Manager">Project Manager</SelectItem>
                      <SelectItem value="Senior Project Manager">Senior Project Manager</SelectItem>
                      <SelectItem value="Project Director">Project Director</SelectItem>
                      <SelectItem value="Engineer">Engineer</SelectItem>
                      <SelectItem value="Architect">Architect</SelectItem>
                      <SelectItem value="Coordinator">Coordinator</SelectItem>
                      <SelectItem value="Specialist">Specialist</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label htmlFor="member-email">Email</Label>
                  <Input
                    id="member-email"
                    type="email"
                    value={newTeamMember.email}
                    onChange={(e) => setNewTeamMember(prev => ({ ...prev, email: e.target.value }))}
                    placeholder="Enter email address"
                  />
                </div>
                <div>
                  <Label htmlFor="member-phone">Phone</Label>
                  <Input
                    id="member-phone"
                    value={newTeamMember.phone}
                    onChange={(e) => setNewTeamMember(prev => ({ ...prev, phone: e.target.value }))}
                    placeholder="Enter phone number"
                  />
                </div>
                <div>
                  <Label htmlFor="member-department">Department</Label>
                  <Input
                    id="member-department"
                    value={newTeamMember.department}
                    onChange={(e) => setNewTeamMember(prev => ({ ...prev, department: e.target.value }))}
                    placeholder="Enter department"
                  />
                </div>
              </div>
              <DialogFooter>
                <Button variant="outline" onClick={() => setShowAddDialog(false)}>
                  Cancel
                </Button>
                <Button onClick={handleAddTeamMember}>
                  Add Team Member
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>
        </div>
      </CardHeader>
      <CardContent>
        <div className="space-y-4">
          {project.team?.map((member) => (
            <div key={member.id} className="flex items-center justify-between p-4 border rounded-lg">
              <div className="flex-1">
                <div className="font-medium">{member.name}</div>
                <div className="text-sm text-gray-600">{member.role}</div>
                <div className="text-sm text-gray-600">{member.department}</div>
                <div className="text-sm text-gray-600">{member.email} • {member.phone}</div>
              </div>
              <Button
                size="sm"
                variant="outline"
                onClick={() => handleRemoveTeamMember(member.id)}
              >
                Remove
              </Button>
            </div>
          ))}
          {(!project.team || project.team.length === 0) && (
            <div className="text-center py-8 text-gray-500">
              No team members added yet
            </div>
          )}
        </div>
      </CardContent>
    </Card>
  )
}

// Project Information Edit Component
function ProjectInfoEdit({ project, onProjectUpdate }) {
  const [showEditDialog, setShowEditDialog] = useState(false)
  const [editedProject, setEditedProject] = useState({
    name: project.name || '',
    description: project.description || '',
    phase: project.phase || '',
    status: project.status || '',
    contractor: project.contractor || '',
    projectManager: project.projectManager || '',
    region: project.region || '',
    startDate: project.startDate || '',
    endDate: project.endDate || '',
    totalBudget: project.totalBudget || 0
  })
  const { showSuccess, showError } = useNotifications()

  const handleSaveChanges = async () => {
    try {
      await onProjectUpdate(project.id, editedProject)
      setShowEditDialog(false)
      showSuccess('Project information updated successfully')
    } catch (error) {
      showError('Failed to update project: ' + error.message)
    }
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div>
            <CardTitle className="flex items-center space-x-2">
              <FileText className="h-5 w-5" />
              <span>Project Information</span>
            </CardTitle>
            <CardDescription>Basic project details and settings</CardDescription>
          </div>
          <Dialog open={showEditDialog} onOpenChange={setShowEditDialog}>
            <DialogTrigger asChild>
              <Button size="sm">
                <Edit className="h-4 w-4 mr-2" />
                Edit Project
              </Button>
            </DialogTrigger>
            <DialogContent className="max-w-2xl">
              <DialogHeader>
                <DialogTitle>Edit Project Information</DialogTitle>
                <DialogDescription>
                  Update the basic information for this project
                </DialogDescription>
              </DialogHeader>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="edit-name">Project Name</Label>
                  <Input
                    id="edit-name"
                    value={editedProject.name}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, name: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="edit-contractor">Contractor</Label>
                  <Input
                    id="edit-contractor"
                    value={editedProject.contractor}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, contractor: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="edit-phase">Phase</Label>
                  <Select value={editedProject.phase} onValueChange={(value) => setEditedProject(prev => ({ ...prev, phase: value }))}>
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Planning">Planning</SelectItem>
                      <SelectItem value="Design">Design</SelectItem>
                      <SelectItem value="Construction">Construction</SelectItem>
                      <SelectItem value="Completion">Completion</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label htmlFor="edit-status">Status</Label>
                  <Select value={editedProject.status} onValueChange={(value) => setEditedProject(prev => ({ ...prev, status: value }))}>
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Active">Active</SelectItem>
                      <SelectItem value="On Hold">On Hold</SelectItem>
                      <SelectItem value="Completed">Completed</SelectItem>
                      <SelectItem value="Cancelled">Cancelled</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label htmlFor="edit-pm">Project Manager</Label>
                  <Input
                    id="edit-pm"
                    value={editedProject.projectManager}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, projectManager: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="edit-region">Region</Label>
                  <Input
                    id="edit-region"
                    value={editedProject.region}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, region: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="edit-start">Start Date</Label>
                  <Input
                    id="edit-start"
                    type="date"
                    value={editedProject.startDate}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, startDate: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="edit-end">End Date</Label>
                  <Input
                    id="edit-end"
                    type="date"
                    value={editedProject.endDate}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, endDate: e.target.value }))}
                  />
                </div>
                <div className="col-span-2">
                  <Label htmlFor="edit-budget">Total Budget</Label>
                  <Input
                    id="edit-budget"
                    type="number"
                    value={editedProject.totalBudget}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, totalBudget: parseFloat(e.target.value) || 0 }))}
                  />
                </div>
                <div className="col-span-2">
                  <Label htmlFor="edit-description">Description</Label>
                  <Textarea
                    id="edit-description"
                    value={editedProject.description}
                    onChange={(e) => setEditedProject(prev => ({ ...prev, description: e.target.value }))}
                    rows={3}
                  />
                </div>
              </div>
              <DialogFooter>
                <Button variant="outline" onClick={() => setShowEditDialog(false)}>
                  Cancel
                </Button>
                <Button onClick={handleSaveChanges}>
                  Save Changes
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>
        </div>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="space-y-4">
            <div>
              <Label className="text-sm font-medium text-gray-600">Project Name</Label>
              <div className="text-lg font-semibold">{project.name}</div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Contractor</Label>
              <div>{project.contractor}</div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Project Manager</Label>
              <div>{project.projectManager}</div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Region</Label>
              <div className="flex items-center">
                <MapPin className="h-4 w-4 mr-2 text-gray-500" />
                {project.region}
              </div>
            </div>
          </div>
          <div className="space-y-4">
            <div>
              <Label className="text-sm font-medium text-gray-600">Phase & Status</Label>
              <div className="flex items-center space-x-2">
                <Badge variant="outline">{project.phase}</Badge>
                <Badge className={getStatusColor(project.status)}>{project.status}</Badge>
              </div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Timeline</Label>
              <div className="flex items-center text-sm">
                <Calendar className="h-4 w-4 mr-2 text-gray-500" />
                {formatDate(project.startDate)} - {formatDate(project.endDate)}
              </div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Budget</Label>
              <div className="flex items-center text-lg font-semibold">
                <DollarSign className="h-4 w-4 mr-1 text-gray-500" />
                {formatCurrency(project.totalBudget)}
              </div>
            </div>
            <div>
              <Label className="text-sm font-medium text-gray-600">Description</Label>
              <div className="text-sm text-gray-700">{project.description}</div>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  )
}

// Enhanced ProjectsPage component
export default function ProjectsPage() {
  const [selectedProject, setSelectedProject] = useState(null)
  const [showPFMTUpload, setShowPFMTUpload] = useState(false)
  const [searchParams, setSearchParams] = useSearchParams()
  const { updateProject } = useProjects()
  const { showSuccess, showError } = useNotifications()

  // Get filter from URL params
  const filterFromUrl = searchParams.get('filter') || 'all'
  
  // Set filter in store when component mounts or URL changes
  useEffect(() => {
    // This will be handled by the ProjectList component
  }, [filterFromUrl])

  const handleProjectUpdate = async (projectId, updates) => {
    try {
      await updateProject(projectId, updates)
      // Update the selected project with new data
      if (selectedProject && selectedProject.id === projectId) {
        setSelectedProject(prev => ({ ...prev, ...updates }))
      }
      showSuccess('Project updated successfully')
    } catch (error) {
      showError('Failed to update project: ' + error.message)
      throw error
    }
  }

  const handlePFMTDataExtracted = async (data) => {
    try {
      // Process the extracted PFMT data and update the project
      const updates = {
        pfmtData: data,
        lastPFMTUpdate: new Date().toISOString()
      }
      await handleProjectUpdate(selectedProject.id, updates)
      setShowPFMTUpload(false)
      showSuccess('PFMT data uploaded and processed successfully')
    } catch (error) {
      showError('Failed to process PFMT data: ' + error.message)
    }
  }

  if (selectedProject) {
    return (
      <div className="p-6">
        <div className="mb-6">
          <Button 
            onClick={() => setSelectedProject(null)}
            variant="outline"
            className="mb-4"
          >
            <ArrowLeft className="h-4 w-4 mr-2" />
            Back to Projects
          </Button>
          
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-3xl font-bold text-gray-900">{selectedProject.name}</h1>
              <p className="text-gray-600 mt-1">
                {selectedProject.contractor} • {selectedProject.phase} • {selectedProject.region}
              </p>
            </div>
            <div className="flex items-center space-x-2">
              <Button 
                onClick={() => setShowPFMTUpload(true)}
                variant="outline"
              >
                <Upload className="h-4 w-4 mr-2" />
                Upload PFMT Data
              </Button>
            </div>
          </div>
        </div>

        {/* Project Summary Cards */}
        <div className="mb-8">
          <ProjectSummary project={selectedProject} />
        </div>

        {/* Project Details Tabs */}
        <Tabs defaultValue="overview" className="space-y-6">
          <TabsList className="grid w-full grid-cols-5">
            <TabsTrigger value="overview">Overview</TabsTrigger>
            <TabsTrigger value="milestones">Milestones</TabsTrigger>
            <TabsTrigger value="vendors">Vendors</TabsTrigger>
            <TabsTrigger value="team">Team</TabsTrigger>
            <TabsTrigger value="documents">Documents</TabsTrigger>
          </TabsList>
          
          <TabsContent value="overview" className="space-y-6">
            <ProjectInfoEdit 
              project={selectedProject} 
              onProjectUpdate={handleProjectUpdate}
            />
          </TabsContent>
          
          <TabsContent value="milestones" className="space-y-6">
            <ProjectMilestones 
              project={selectedProject} 
              onProjectUpdate={handleProjectUpdate}
            />
          </TabsContent>
          
          <TabsContent value="vendors" className="space-y-6">
            <ProjectVendors 
              project={selectedProject} 
              onProjectUpdate={handleProjectUpdate}
            />
          </TabsContent>
          
          <TabsContent value="team" className="space-y-6">
            <ProjectTeam 
              project={selectedProject} 
              onProjectUpdate={handleProjectUpdate}
            />
          </TabsContent>
          
          <TabsContent value="documents" className="space-y-6">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <FileText className="h-5 w-5" />
                  <span>Project Documents</span>
                </CardTitle>
                <CardDescription>Manage project documents and PFMT data</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <Button 
                    onClick={() => setShowPFMTUpload(true)}
                    className="w-full"
                  >
                    <Upload className="h-4 w-4 mr-2" />
                    Upload PFMT Excel File
                  </Button>
                  
                  {selectedProject.pfmtData && (
                    <div className="p-4 bg-green-50 border border-green-200 rounded-lg">
                      <div className="flex items-center space-x-2">
                        <CheckCircle className="h-5 w-5 text-green-600" />
                        <span className="font-medium text-green-800">PFMT Data Available</span>
                      </div>
                      <div className="text-sm text-green-700 mt-1">
                        Last updated: {formatDate(selectedProject.lastPFMTUpdate)}
                      </div>
                    </div>
                  )}
                  
                  <div className="text-center py-8 text-gray-500">
                    Document management features coming soon
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>

        {/* PFMT Upload Dialog */}
        {showPFMTUpload && (
          <Dialog open={showPFMTUpload} onOpenChange={setShowPFMTUpload}>
            <DialogContent className="max-w-4xl">
              <DialogHeader>
                <DialogTitle>Upload PFMT Data</DialogTitle>
                <DialogDescription>
                  Upload and extract data from PFMT Excel file for {selectedProject.name}
                </DialogDescription>
              </DialogHeader>
              <PFMTDataExtractor
                project={selectedProject}
                onDataExtracted={handlePFMTDataExtracted}
                onClose={() => setShowPFMTUpload(false)}
              />
            </DialogContent>
          </Dialog>
        )}
      </div>
    )
  }
  
  return (
    <div className="p-6">
      <div className="mb-6">
        <h1 className="text-3xl font-bold text-gray-900 mb-2">
          {filterFromUrl === 'my' ? 'My Projects' : 'All Projects'}
        </h1>
        <p className="text-gray-600">
          {filterFromUrl === 'my' 
            ? 'Manage and track your assigned projects'
            : 'Manage and track all projects in the system'
          }
        </p>
        
        {/* Filter indicator */}
        {filterFromUrl !== 'all' && (
          <div className="mt-4">
            <Badge variant="outline" className="flex items-center w-fit">
              <Filter className="h-3 w-3 mr-1" />
              Filter: {filterFromUrl === 'my' ? 'My Projects' : filterFromUrl}
            </Badge>
          </div>
        )}
      </div>
      
      <ProjectList onProjectSelect={setSelectedProject} filter={filterFromUrl} />
    </div>
  )
}

