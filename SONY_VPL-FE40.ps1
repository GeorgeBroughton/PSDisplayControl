class SONY_VPL_FE40 {
    
    hidden [System.IO.Ports.SerialPort]$Device = [System.IO.Ports.SerialPort]::New()
    [string]$Port = 'COM1'
    [int]$MonitorID = 1

    [string]$SICPVersion
    [string]$PlatformLabel
    [string]$PlatformVersion
    
    [String]$ModelNumber
    [string]$FirmwareVersion
    [string]$BuildDate

    hidden [HashTable]$TranslationTable = @{
        # put lookup tables here
    }

    [void] Connect () {
        if ($This.Device) {$This.Device.close()}
        try {
            $This.Device.PortName       = $This.Port
            $This.Device.BaudRate       = 9600
            $This.Device.DataBits       = 8
            $This.Device.Parity         = "None"
            $This.Device.StopBits       = 1
            $This.Device.Handshake      = "None"
            $This.Device.ReadTimeout    = 1000
            $This.Device.WriteTimeout   = 1000
            $This.Device.open()
            
        }
        Catch {
            $This.Device.close()
            Continue
        }
        return
    }

    [void] Disconnect () {
        $This.Device.close()
        return
    }

    # Set Commands

    # Get Commands

    # For processing commands

    hidden [PSCustomObject] Receive () {
        return @{}
    }

    hidden [void] Send ([byte[]]$Message) {
        return
    }

}