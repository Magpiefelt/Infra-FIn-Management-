// Per-project data management components
import React, { useState } from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog.jsx'
import { Input } from '@/components/ui/input.jsx'
import { Label } from '@/components/ui/label.jsx'
import { Textarea } from '@/components/ui/textarea.jsx'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select.jsx'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Plus, Edit, Trash2, Users, Building, CheckCircle } from 'lucide-react'
import { formatDate, getStatusColor } from '../../utils/index.js'
import { useNotifications } from '../../hooks/index.js'

export function ProjectMilestones({ project, onProjectUpdate }) {
  const [editingMilestone, setEditingMilestone] = useState(null)
  const [showAddDialog, setShowAddDialog] = useState(false)
  const [newMilestone, setNewMilestone] = useState({
    name: '',
    description: '',
    plannedDate: '',
    status: 'pending'
  })
  const { showSuccess, showError } = useNotifications()

  const handleUpdateMilestone = async (milestoneId, updates) => {
    try {
      const updatedMilestones = project.milestones.map(milestone =>
        milestone.id === milestoneId ? { ...milestone, ...updates } : milestone
      )
      
      await onProjectUpdate(project.id, { milestones: updatedMilestones })
      setEditingMilestone(null)
      showSuccess('Milestone updated successfully')
    } catch (error) {
      showError('Failed to update milestone: ' + error.message)
    }
  }

  const handleAddMilestone = async () => {
    try {
      const milestone = {
        id: Date.now(),
        ...newMilestone,
        actualDate: null,
        baselineDate: newMilestone.plannedDate,
        notes: '',
        isNA: false,
        isRequired: false
      }
      
      const updatedMilestones = [...project.milestones, milestone]
      await onProjectUpdate(project.id, { milestones: updatedMilestones })
      
      setNewMilestone({ name: '', description: '', plannedDate: '', status: 'pending' })
      setShowAddDialog(false)
      showSuccess('Milestone added successfully')
    } catch (error) {
      showError('Failed to add milestone: ' + error.message)
    }
  }

  const handleDeleteMilestone = async (milestoneId) => {
    try {
      const updatedMilestones = project.milestones.filter(m => m.id !== milestoneId)
      await onProjectUpdate(project.id, { milestones: updatedMilestones })
      showSuccess('Milestone deleted successfully')
    } catch (error) {
      showError('Failed to delete milestone: ' + error.message)
    }
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div>
            <CardTitle className="flex items-center space-x-2">
              <CheckCircle className="h-5 w-5" />
              <span>Project Milestones</span>
            </CardTitle>
            <CardDescription>Track project milestones and key deliverables</CardDescription>
          </div>
          <Dialog open={showAddDialog} onOpenChange={setShowAddDialog}>
            <DialogTrigger asChild>
              <Button size="sm">
                <Plus className="h-4 w-4 mr-2" />
                Add Milestone
              </Button>
            </DialogTrigger>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Add New Milestone</DialogTitle>
                <DialogDescription>
                  Create a new milestone for this project
                </DialogDescription>
              </DialogHeader>
              <div className="space-y-4">
                <div>
                  <Label htmlFor="milestone-name">Milestone Name</Label>
                  <Input
                    id="milestone-name"
                    value={newMilestone.name}
                    onChange={(e) => setNewMilestone(prev => ({ ...prev, name: e.target.value }))}
                    placeholder="Enter milestone name"
                  />
                </div>
                <div>
                  <Label htmlFor="milestone-description">Description</Label>
                  <Textarea
                    id="milestone-description"
                    value={newMilestone.description}
                    onChange={(e) => setNewMilestone(prev => ({ ...prev, description: e.target.value }))}
                    placeholder="Enter milestone description"
                  />
                </div>
                <div>
                  <Label htmlFor="milestone-date">Planned Date</Label>
                  <Input
                    id="milestone-date"
                    type="date"
                    value={newMilestone.plannedDate}
                    onChange={(e) => setNewMilestone(prev => ({ ...prev, plannedDate: e.target.value }))}
                  />
                </div>
                <div>
                  <Label htmlFor="milestone-status">Status</Label>
                  <Select value={newMilestone.status} onValueChange={(value) => setNewMilestone(prev => ({ ...prev, status: value }))}>
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="pending">Pending</SelectItem>
                      <SelectItem value="in-progress">In Progress</SelectItem>
                      <SelectItem value="completed">Completed</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>
              <DialogFooter>
                <Button variant="outline" onClick={() => setShowAddDialog(false)}>
                  Cancel
                </Button>
                <Button onClick={handleAddMilestone}>
                  Add Milestone
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>
        </div>
      </CardHeader>
      <CardContent>
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Milestone</TableHead>
              <TableHead>Status</TableHead>
              <TableHead>Planned Date</TableHead>
              <TableHead>Actual Date</TableHead>
              <TableHead>Actions</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {project.milestones?.map((milestone) => (
              <TableRow key={milestone.id}>
                <TableCell>
                  <div>
                    <div className="font-medium">{milestone.name}</div>
                    <div className="text-sm text-gray-600">{milestone.description}</div>
                  </div>
                </TableCell>
                <TableCell>
                  <Badge className={getStatusColor(milestone.status)}>
                    {milestone.status}
                  </Badge>
                </TableCell>
                <TableCell>{formatDate(milestone.plannedDate)}</TableCell>
                <TableCell>{formatDate(milestone.actualDate)}</TableCell>
                <TableCell>
                  <div className="flex space-x-2">
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => setEditingMilestone(milestone)}
                    >
                      <Edit className="h-4 w-4" />
                    </Button>
                    {!milestone.isRequired && (
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => handleDeleteMilestone(milestone.id)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  )
}

export function ProjectVendors({ project, onProjectUpdate }) {
  const [showAddDialog, setShowAddDialog] = useState(false)
  const [newVendor, setNewVendor] = useState({
    name: '',
    type: '',
    contact: '',
    email: '',
    phone: '',
    status: 'Active'
  })
  const { showSuccess, showError } = useNotifications()

  const handleAddVendor = async () => {
    try {
      const vendor = {
        id: Date.now(),
        ...newVendor
      }
      
      const updatedVendors = [...(project.vendors || []), vendor]
      await onProjectUpdate(project.id, { vendors: updatedVendors })
      
      setNewVendor({ name: '', type: '', contact: '', email: '', phone: '', status: 'Active' })
      setShowAddDialog(false)
      showSuccess('Vendor added successfully')
    } catch (error) {
      showError('Failed to add vendor: ' + error.message)
    }
  }

  const handleDeleteVendor = async (vendorId) => {
    try {
      const updatedVendors = project.vendors.filter(v => v.id !== vendorId)
      await onProjectUpdate(project.id, { vendors: updatedVendors })
      showSuccess('Vendor removed successfully')
    } catch (error) {
      showError('Failed to remove vendor: ' + error.message)
    }
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div>
            <CardTitle className="flex items-center space-x-2">
              <Building className="h-5 w-5" />
              <span>Project Vendors</span>
            </CardTitle>
            <CardDescription>Manage vendors and contractors for this project</CardDescription>
          </div>
          <Dialog open={showAddDialog} onOpenChange={setShowAddDialog}>
            <DialogTrigger asChild>
              <Button size="sm">
                <Plus className="h-4 w-4 mr-2" />
                Add Vendor
              </Button>
            </DialogTrigger>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Add New Vendor</DialogTitle>
                <DialogDescription>
                  Add a vendor or contractor to this project
                </DialogDescription>
              </DialogHeader>
              <div className="space-y-4">
                <div>
                  <Label htmlFor="vendor-name">Vendor Name</Label>
                  <Input
                    id="vendor-name"
                    value={newVendor.name}
                    onChange={(e) => setNewVendor(prev => ({ ...prev, name: e.target.value }))}
                    placeholder="Enter vendor name"
                  />
                </div>
                <div>
                  <Label htmlFor="vendor-type">Type</Label>
                  <Select value={newVendor.type} onValueChange={(value) => setNewVendor(prev => ({ ...prev, type: value }))}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select vendor type" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="General Contractor">General Contractor</SelectItem>
                      <SelectItem value="Subcontractor">Subcontractor</SelectItem>
                      <SelectItem value="Material Supplier">Material Supplier</SelectItem>
                      <SelectItem value="Consultant">Consultant</SelectItem>
                      <SelectItem value="Equipment Rental">Equipment Rental</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label htmlFor="vendor-contact">Contact Person</Label>
                  <Input
                    id="vendor-contact"
                    value={newVendor.contact}
                    onChange={(e) => setNewVendor(prev => ({ ...prev, contact: e.target.value }))}
                    placeholder="Enter contact person name"
                  />
                </div>
                <div>
                  <Label htmlFor="vendor-email">Email</Label>
                  <Input
                    id="vendor-email"
                    type="email"
                    value={newVendor.email}
                    onChange={(e) => setNewVendor(prev => ({ ...prev, email: e.target.value }))}
                    placeholder="Enter email address"
                  />
                </div>
                <div>
                  <Label htmlFor="vendor-phone">Phone</Label>
                  <Input
                    id="vendor-phone"
                    value={newVendor.phone}
                    onChange={(e) => setNewVendor(prev => ({ ...prev, phone: e.target.value }))}
                    placeholder="Enter phone number"
                  />
                </div>
              </div>
              <DialogFooter>
                <Button variant="outline" onClick={() => setShowAddDialog(false)}>
                  Cancel
                </Button>
                <Button onClick={handleAddVendor}>
                  Add Vendor
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>
        </div>
      </CardHeader>
      <CardContent>
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Vendor Name</TableHead>
              <TableHead>Type</TableHead>
              <TableHead>Contact</TableHead>
              <TableHead>Status</TableHead>
              <TableHead>Actions</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {project.vendors?.map((vendor) => (
              <TableRow key={vendor.id}>
                <TableCell className="font-medium">{vendor.name}</TableCell>
                <TableCell>{vendor.type}</TableCell>
                <TableCell>
                  <div>
                    <div>{vendor.contact}</div>
                    <div className="text-sm text-gray-600">{vendor.email}</div>
                    <div className="text-sm text-gray-600">{vendor.phone}</div>
                  </div>
                </TableCell>
                <TableCell>
                  <Badge variant={vendor.status === 'Active' ? 'default' : 'secondary'}>
                    {vendor.status}
                  </Badge>
                </TableCell>
                <TableCell>
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => handleDeleteVendor(vendor.id)}
                  >
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  )
}

