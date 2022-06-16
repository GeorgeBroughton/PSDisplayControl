class iiyama {
    
    hidden [System.IO.Ports.SerialPort]$Device = [System.IO.Ports.SerialPort]::New()
    [string]$Port = 'COM1'
    [int]$MonitorID = 1

    hidden [HashTable]$TranslationTable = @{
        Power = [PSCustomObject]@{
            'Off'                       = 0x01
            'On'                        = 0x02
        }
        ReturnPower = [PSCustomObject]@{
            'Off'                       = 0x01
            'On'                        = 0x02
            'Last State'                = 0x03
        }
        Source = [PSCustomObject]@{
            'VIDEO'                     = 0x01
            'S-VIDEO'                   = 0x02
            'COMPONENT'                 = 0x03
            'CVI2'                      = 0x04
            'VGA'                       = 0x05
            'HDMI2'                     = 0x06
            'DisplayPort2'              = 0x07
            'USB2'                      = 0x08
            'Card DVI-D'                = 0x09
            'DisplayPort1'              = 0x0A
            'Card OPS'                  = 0x0B
            'USB1'                      = 0x0C
            'HDMI'                      = 0x0D
            'DVI-D'                     = 0x0E
            'HDMI3'                     = 0x0F
            'BROWSER'                   = 0x10
            'SMARTCMS'                  = 0x11
            'DMS'                       = 0x12
            'INTERNAL STORAGE'          = 0x13
            'Reserved1'                 = 0x14
            'Reserved2'                 = 0x15
            'Media Player'              = 0x16
            'PDF Player'                = 0x17
            'Custom'                    = 0x18
            'HDMI4'                     = 0x19
        }
        Gamma = [PSCustomObject]@{
            'Native'                    = 0x01
            'S-Gamma'                   = 0x02
            '2.2'                       = 0x03
            '2.4'                       = 0x04
            'D-Image'                   = 0x05
        }
        ColorTemp = [PSCustomObject]@{
            'User1'                     = 0x00
            'Native'                    = 0x01
            '11000K'                    = 0x02
            '10000K'                    = 0x03
            '9300K'                     = 0x04
            '7500K'                     = 0x05
            '6500K'                     = 0x06
            '5770K'                     = 0x07
            '5500K'                     = 0x08
            '5000K'                     = 0x09
            '4000K'                     = 0x0A
            '3400K'                     = 0x0B
            '3550K'                     = 0x0C
            '3000K'                     = 0x0D
            '2800K'                     = 0x0E
            '2600K'                     = 0x0F
            '1850K'                     = 0x10
            'User2'                     = 0x12
        }
        PictureFormat = [PSCustomObject]@{
            '4:3'                       = 0x00
            'Custom'                    = 0x01
            '1:1'                       = 0x02  
            'Full'                      = 0x03
            '21:9'                      = 0x04
            'Dynamic'                   = 0x05
            '16:9'                      = 0x06
        }
        KeypadLock = [PSCustomObject]@{
            'Unlocked'                  = 0x01
            'Locked'                    = 0x02
            'Power Unlocked'            = 0x03
            'Volume Unlocked'           = 0x04
            'Power & Volume Unlocked'   = 0x07
        }
        IRRemoteLock = [PSCustomObject]@{
            'Unlocked'                  = 0x01
            'Locked'                    = 0x02
            'Power Unlocked'            = 0x03
            'Volume Unlocked'           = 0x04
            'Primary'                   = 0x05
            'Secondary'                 = 0x06
            'Power & Volume Unlocked'   = 0x07
        }
        CommunicationControl = [PSCustomObject]@{
            'Completed'                 = 0x00
            'Excessive Parameters'      = 0x01
            'Insufficient Parameters'   = 0x02
            'Cancelled'                 = 0x03
            'Error'                     = 0x04
        }
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

    [PSCustomObject] SetPower ([string]$State) {
        if ( $This.TranslationTable.Power.$State ) {
            [byte]$Command = $This.TranslationTable.Power.$State
            $This.Send((0x18,$Command))
            if ($State -eq 'On') {
                return [PSCustomObject]@{ Status = "Command Sent" }
            }
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetSource ([string]$Source) {
        if ( $This.TranslationTable.Source.$Source ) {
            [byte]$Command = $This.TranslationTable.Source.$Source
            $This.Send((0xAC,$Command,0x00,0x00,0x00))
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetVideoParameters ([int]$Brightness,[int]$Color,[int]$Contrast,[int]$Sharpness,[int]$Tint,[int]$BlackLevel,[string]$Gamma) {
        if ( $This.TranslationTable.Gamma.$Gamma ) {
            $This.Send(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,$This.TranslationTable.Gamma.$Gamma)
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetColorTemperature ([string]$Temperature) {
        if ( $This.TranslationTable.ColorTemp.$Temperature ) {
            $This.Send((0x34,$This.TranslationTable.ColorTemp.$Temperature))
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetColorParameters ([int]$Red,[int]$Green,[int]$Blue,[int]$RedOffset,[int]$GreenOffset,[int]$BlueOffset) {
        $This.Send(0x36,$Red,$Green,$Blue,$RedOffset,$GreenOffset,$BlueOffset)
        return $This.Receive()
    }

    [PSCustomObject] SetPictureFormat ([string]$PictureFormat) {
        if ( $This.TranslationTable.PictureFormat.$PictureFormat ) {
            $This.Send((0x3A,$This.TranslationTable.PictureFormat.$PictureFormat))
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetIRLockStatus ([string]$LockStatus) {
        if ( $This.TranslationTable.IRRemoteLock.$LockStatus ) {
            $This.Send((0x1C,$This.TranslationTable.IRRemoteLock.$LockStatus))
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetKeypadLockStatus ([string]$LockStatus) {
        if ( $This.TranslationTable.KeypadLock.$LockStatus ) {
            $This.Send((0x1A,$This.TranslationTable.KeypadLock.$LockStatus))
            return $This.Receive()
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetVolume ([int]$Volume,[int]$AudioOutVolume) {
        $This.Send((0x44,$Volume,$AudioOutVolume))
        return $This.Receive()
    }

    # Get Commands



    [PSCustomObject] GetPower () {
        $This.Send((0x19))
        return $This.Receive()
    }

    [PSCustomObject] GetIRLockStatus () {
        $This.Send((0x1D))
        return $This.Receive()
    }

    [PSCustomObject] GetKeypadLockStatus () {
        $This.Send((0x1B))
        return $This.Receive()
    }

    [PSCustomObject] GetReturnPowerState () {
        $This.Send((0xA4))
        return $This.Receive()
    }

    [PSCustomObject] GetSource () {
        $This.Send((0xAD))
        return $This.Receive()
    }

    [PSCustomObject] GetVideoParameters () {
        $This.Send((0x33))
        return $This.Receive()
    }

    [PSCustomObject] GetColorTemperature () {
        $This.Send((0x35))
        return $This.Receive()
    }

    [PSCustomObject] GetColorParameters () {
        $This.Send((0x37))
        return $This.Receive()
    }

    [PSCustomObject] GetPictureFormat () {
        $This.Send((0x3B))
        return $This.Receive()
    }

    [PSCustomObject] GetVolume () {
        $This.Send((0x45))
        return $This.Receive()
    }

    [PSCustomObject] GetAudioParameters () {
        $This.Send((0x43))
        return $This.Receive()
    }

    [PSCustomObject] GetOperatingHours () {
        $This.Send((0x0F))
        return $This.Receive()
    }

    # For processing commands

    hidden [PSCustomObject] Receive () {
        [byte[]]$ReadBuffer = @()

        foreach ($i in 0..4)                { $ReadBuffer += $This.Device.ReadByte() }
        foreach ($i in 1..$ReadBuffer[4])   { $ReadBuffer += $This.Device.ReadByte() }
        
        if ($ReadBuffer[0] -eq 0x21) {

            if ($ReadBuffer[1] -eq $This.MonitorID) {
                
                if ($ReadBuffer[2] -eq 0x00) {

                    if ($ReadBuffer[3] -eq 0x00) {

                        [byte]$XOR = 0x00                        
                        foreach ($i in $ReadBuffer[0..($ReadBuffer.count - 2)]) {
                            [byte]$XOR = $XOR -bxor $i
                        }

                        if ($ReadBuffer[-1] -ne $XOR) { return [PSCustomObject]@{ Status = 'Error' ; Details = "Bad Checksum." } }

                        switch ($ReadBuffer[6]) {
                            0x00 {
                                $Value = $This.TranslationTable.CommunicationControl.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Status = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Communication Control: Property out of bounds." }
                            }
                            0x19 {
                                $Value = $This.TranslationTable.Power.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Power_State = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Power State: Property out of bounds." }
                            }
                            0x1D {
                                $Value = $This.TranslationTable.IRRemoteLock.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{IR_Remote_Lock = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "IR Lock Status: Property out of bounds." }
                            }
                            0x1B {
                                $Value = $This.TranslationTable.KeypadLock.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Keypad_Lock = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Keypad Lock Status: Property out of bounds." }
                            }
                            0xA2 {
                                return [PSCustomObject]@{Response = $ReadBuffer[7..($ReadBuffer.Count-2)] }
                            }
                            0xA4 {
                                $Value = $This.TranslationTable.ReturnPower.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Return_Power_State = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Return Power Status: Property out of bounds." }
                            }
                            0xAD {
                                $Value = $This.TranslationTable.Source.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Current_Source = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Current Source: Property out of bounds." }
                            }
                            0x33 {
                                $Value = $This.TranslationTable.Gamma.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) {
                                    if (( $ReadBuffer[7] -le 100 ) -and ( $ReadBuffer[7] -ge 0 )){
                                        if (( $ReadBuffer[8] -le 100 ) -and ( $ReadBuffer[8] -ge 0 )) {
                                            if (( $ReadBuffer[9] -le 100 ) -and ( $ReadBuffer[9] -ge 0 )) {
                                                if (( $ReadBuffer[10] -le 100 ) -and ( $ReadBuffer[10] -ge 0 )) {
                                                    if (( $ReadBuffer[11] -le 100 ) -and ( $ReadBuffer[11] -ge 0 )) {
                                                        if (( $ReadBuffer[12] -le 100 ) -and ( $ReadBuffer[12] -ge 0 )) {
                                                            if ( 0x01,0x02,0x03,0x04,0x05 -contains $ReadBuffer[13] ) {
                                                                return [PSCustomObject]@{
                                                                    Brightness  = $ReadBuffer[7]
                                                                    Color       = $ReadBuffer[8]
                                                                    Contrast    = $ReadBuffer[9]
                                                                    Sharpness   = $ReadBuffer[10]
                                                                    Tint        = $ReadBuffer[11]
                                                                    BlackLevel  = $ReadBuffer[12]
                                                                    Gamma       = $Value
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Video Parameters: Property out of bounds." }
                            }
                            0x35 {
                                $Value = $This.TranslationTable.ColorTemp.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Color_Temperature = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Color Temperature: Property out of bounds." }
                            }
                            0x37 {
                                return [PSCustomObject]@{
                                    Red_Gain     = $ReadBuffer[7]
                                    Green_Gain   = $ReadBuffer[8]
                                    Blue_Gain    = $ReadBuffer[9]
                                    Red_Offset   = $ReadBuffer[10]
                                    Green_Offset = $ReadBuffer[11]
                                    Blue_Offset  = $ReadBuffer[12]
                                }
                            }
                            0x3B {
                                $Value = $This.TranslationTable.PictureFormat.PSObject.Properties.Where({$_.Value -eq $ReadBuffer[7]}).Name
                                if ($Value) { return [PSCustomObject]@{Picture_Format = $Value } }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Picture Format: Property out of bounds." }
                            }
                            0x45 {
                                if (( $ReadBuffer[7] -le 100 ) -and ( $ReadBuffer[7] -ge 0 )){
                                    if (( $ReadBuffer[8] -le 100 ) -and ( $ReadBuffer[8] -ge 0 )) {
                                        return [PSCustomObject]@{
                                            Volume         = $ReadBuffer[7]
                                            Audio_Output_Volume = $ReadBuffer[8]
                                        }
                                    }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Volume: Property out of bounds." }
                            }
                            0x43 {
                                if (( $ReadBuffer[7] -le 100 ) -and ( $ReadBuffer[7] -ge 0 )){
                                    if (( $ReadBuffer[8] -le 100 ) -and ( $ReadBuffer[8] -ge 0 )) {
                                        return [PSCustomObject]@{
                                            Treble = $ReadBuffer[7]
                                            Bass   = $ReadBuffer[8]
                                        }
                                    }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Audio Parameters: Property out of bounds." }
                            }
                            0x0F {
                                [PSCustomObject]@{
                                    Operating_Hours = [System.BitConverter]::ToInt16($ReadBuffer,2)
                                }
                            }
                        }
                    }
                }
            }
        }
        return [PSCustomObject]@{ Status = 'Error' ; Details = "Response contained garbage." ; Buffer = $ReadBuffer }
        $This.Device.DiscardInBuffer()
    }

    hidden [void] Send ([byte[]]$Message) {

        [byte[]]$SendBuffer = 0xA6,$This.MonitorID,0x00,0x00,0x00

        $SendBuffer += [byte]$($Message.Length + 2)
        $SendBuffer += [byte]0x01
        $SendBuffer += $Message

        [byte]$XOR = 0x00
        foreach ($i in $SendBuffer) {
            [byte]$XOR = $XOR -bxor $i
        }
        
        $SendBuffer += $XOR

        $This.Device.Write($SendBuffer,0,$SendBuffer.Count)
        
    }

}