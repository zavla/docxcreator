// works with Windows Service Control Manager
package main

import "golang.org/x/sys/windows/svc"

// Tservice represents my service and has a method Execute
type Tservice struct {
	currentConfig serviceConfig
}

// Execute responds to SCM
func (s *Tservice) Execute(args []string, changerequest <-chan svc.ChangeRequest, updatestatus chan<- svc.Status) (ssec bool, errno uint32) {
	updatestatus <- svc.Status{State: svc.StartPending}

	//go runHTTP(s.currentConfig.bindAddressPort)

	supports := svc.AcceptStop | svc.AcceptShutdown

	updatestatus <- svc.Status{State: svc.Running, Accepts: supports}
	// select has no default and wait indefinitly
	select {
	case c := <-changerequest:
		switch c.Cmd {
		case svc.Stop, svc.Shutdown:
			goto stoped
		case svc.Interrogate:

		}
	}
stoped:
	updatestatus <- svc.Status{State: svc.StopPending}
	return false, 0
}
