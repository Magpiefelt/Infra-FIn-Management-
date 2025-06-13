// Workflow management components
import React, { useState } from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Alert, AlertDescription } from '@/components/ui/alert.jsx'
import { CheckCircle, Archive, RotateCcw, AlertTriangle } from 'lucide-react'
import { formatDate, formatDateTime } from '../../utils/index.js'
import { useNotifications } from '../../hooks/index.js'

export function WorkflowActions({ project, onProjectUpdate }) {
  const [showArchiveDialog, setShowArchiveDialog] = useState(false)
  const [archiving, setArchiving] = useState(false)
  const { showSuccess, showError } = useNotifications()

  const handleDirectorApproval = async () => {
    try {
      const updates = {
        directorApproved: true,
        approvedBy: "Lisa Wilson", // Current director
        approvedDate: new Date().toISOString(),
        reportStatus: "Approved"
      }
      
      await onProjectUpdate(project.id, updates)
      showSuccess('Project approved by Director')
      
      // Show archive dialog after approval
      setShowArchiveDialog(true)
    } catch (error) {
      showError('Failed to approve project: ' + error.message)
    }
  }

  const handleArchiveAndReset = async () => {
    setArchiving(true)
    try {
      // Import the archiving function
      const { archiveMonthlyData } = await import('../../services/mockData.js')
      
      const result = await archiveMonthlyData(project.id)
      
      // Update the project with reset status
      await onProjectUpdate(project.id, {
        reportStatus: "Update Required",
        scheduleReasonCode: "",
        budgetReasonCode: "",
        monthlyComments: "",
        previousHighlights: "",
        nextSteps: "",
        budgetVarianceExplanation: "",
        cashflowVarianceExplanation: "",
        submittedBy: null,
        submittedDate: null,
        approvedBy: null,
        approvedDate: null,
        directorApproved: false,
        seniorPmReviewed: false
      })
      
      showSuccess(result.message)
      setShowArchiveDialog(false)
    } catch (error) {
      showError('Failed to archive data: ' + error.message)
    } finally {
      setArchiving(false)
    }
  }

  const canApprove = project.submittedBy && !project.directorApproved
  const isApproved = project.directorApproved

  return (
    <>
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center space-x-2">
            <CheckCircle className="h-5 w-5" />
            <span>Workflow Actions</span>
          </CardTitle>
          <CardDescription>
            Manage project approval workflow and monthly cycles
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="flex items-center justify-between p-4 border rounded-lg">
            <div>
              <h4 className="font-medium">Current Status</h4>
              <p className="text-sm text-gray-600">
                <Badge variant={project.reportStatus === 'Approved' ? 'default' : 'destructive'}>
                  {project.reportStatus}
                </Badge>
              </p>
            </div>
            
            {project.submittedBy && (
              <div className="text-right">
                <p className="text-sm font-medium">Submitted by {project.submittedBy}</p>
                <p className="text-xs text-gray-600">{formatDateTime(project.submittedDate)}</p>
              </div>
            )}
          </div>

          {isApproved && (
            <Alert>
              <CheckCircle className="h-4 w-4" />
              <AlertDescription>
                Project approved by {project.approvedBy} on {formatDate(project.approvedDate)}
              </AlertDescription>
            </Alert>
          )}

          <div className="flex space-x-2">
            {canApprove && (
              <Button onClick={handleDirectorApproval} className="flex-1">
                <CheckCircle className="h-4 w-4 mr-2" />
                Director Approval
              </Button>
            )}
            
            {isApproved && (
              <Button 
                onClick={() => setShowArchiveDialog(true)}
                variant="outline" 
                className="flex-1"
              >
                <Archive className="h-4 w-4 mr-2" />
                Archive & Reset for Next Month
              </Button>
            )}
          </div>

          {project.monthlyArchive && project.monthlyArchive.length > 0 && (
            <div className="mt-4">
              <h4 className="font-medium mb-2">Monthly Archive History</h4>
              <div className="space-y-2 max-h-40 overflow-y-auto">
                {project.monthlyArchive.map((archive, index) => (
                  <div key={archive.id} className="flex items-center justify-between p-2 bg-gray-50 rounded text-sm">
                    <span>{archive.reportingPeriod}</span>
                    <span className="text-gray-600">{formatDate(archive.archivedDate)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Archive Confirmation Dialog */}
      <Dialog open={showArchiveDialog} onOpenChange={setShowArchiveDialog}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle className="flex items-center space-x-2">
              <Archive className="h-5 w-5" />
              <span>Archive Monthly Data</span>
            </DialogTitle>
            <DialogDescription>
              This will archive the current month's data and reset the project status for the next reporting cycle.
            </DialogDescription>
          </DialogHeader>
          
          <div className="py-4">
            <Alert>
              <AlertTriangle className="h-4 w-4" />
              <AlertDescription>
                <strong>This action will:</strong>
                <ul className="list-disc list-inside mt-2 space-y-1">
                  <li>Save current month's data to archive</li>
                  <li>Reset project status to "Update Required"</li>
                  <li>Clear monthly comments and variance explanations</li>
                  <li>Reset approval workflow for next cycle</li>
                </ul>
              </AlertDescription>
            </Alert>
          </div>

          <DialogFooter>
            <Button 
              variant="outline" 
              onClick={() => setShowArchiveDialog(false)}
              disabled={archiving}
            >
              Cancel
            </Button>
            <Button 
              onClick={handleArchiveAndReset}
              disabled={archiving}
            >
              {archiving ? (
                <>
                  <RotateCcw className="h-4 w-4 mr-2 animate-spin" />
                  Archiving...
                </>
              ) : (
                <>
                  <Archive className="h-4 w-4 mr-2" />
                  Archive & Reset
                </>
              )}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </>
  )
}

