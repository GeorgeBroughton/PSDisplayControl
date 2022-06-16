class iiyama {
    
    hidden [System.IO.Ports.SerialPort]$Device = [System.IO.Ports.SerialPort]::New()
    [string]$Port = 'COM1'
    [int]$MonitorID = 1

    [void] Connect () {
        if ($This.Device) {$This.Device.close()}
        try {
            $This.Device.PortName = $This.Port
            $This.Device.BaudRate = 9600
            $This.Device.DataBits = 8
            $This.Device.Parity = "None"
            $This.Device.StopBits = 1
            $This.Device.Handshake = "None"
            $This.Device.ReadTimeout = 1000
            $This.Device.WriteTimeout = 1000
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

    [PSCustomObject] SetPower ([string]$State) {
        switch ($State) {
            'Off'  { return $This.Command((0x18,0x01))}
            'On' { return $This.Command((0x18,0x02))}
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetSource ([string]$State) {
        switch ($State) {
            'VIDEO'             { return $This.Command((0xAC,0x01,0x00,0x00,0x00))}
            'S-VIDEO'           { return $This.Command((0xAC,0x02,0x00,0x00,0x00))}
            'COMPONENT'         { return $This.Command((0xAC,0x03,0x00,0x00,0x00))}
            'CVI2'              { return $This.Command((0xAC,0x04,0x00,0x00,0x00))}
            'VGA'               { return $This.Command((0xAC,0x05,0x00,0x00,0x00))}
            'HDMI2'             { return $This.Command((0xAC,0x06,0x00,0x00,0x00))}
            'DisplayPort2'      { return $This.Command((0xAC,0x07,0x00,0x00,0x00))}
            'USB2'              { return $This.Command((0xAC,0x08,0x00,0x00,0x00))}
            'Card DVI-D'        { return $This.Command((0xAC,0x09,0x00,0x00,0x00))}
            'DisplayPort1'      { return $This.Command((0xAC,0x0A,0x00,0x00,0x00))}
            'Card OPS'          { return $This.Command((0xAC,0x0B,0x00,0x00,0x00))}
            'USB1'              { return $This.Command((0xAC,0x0C,0x00,0x00,0x00))}
            'HDMI'              { return $This.Command((0xAC,0x0D,0x00,0x00,0x00))}
            'DVI-D'             { return $This.Command((0xAC,0x0E,0x00,0x00,0x00))}
            'HDMI3'             { return $This.Command((0xAC,0x0F,0x00,0x00,0x00))}
            'BROWSER'           { return $This.Command((0xAC,0x10,0x00,0x00,0x00))}
            'SMARTCMS'          { return $This.Command((0xAC,0x11,0x00,0x00,0x00))}
            'DMS'               { return $This.Command((0xAC,0x12,0x00,0x00,0x00))}
            'INTERNAL STORAGE'  { return $This.Command((0xAC,0x13,0x00,0x00,0x00))}
            'Reserved1'         { return $This.Command((0xAC,0x14,0x00,0x00,0x00))}
            'Reserved2'         { return $This.Command((0xAC,0x15,0x00,0x00,0x00))}
            'Media Player'      { return $This.Command((0xAC,0x16,0x00,0x00,0x00))}
            'PDF Player'        { return $This.Command((0xAC,0x17,0x00,0x00,0x00))}
            'Custom'            { return $This.Command((0xAC,0x18,0x00,0x00,0x00))}
            'HDMI4'             { return $This.Command((0xAC,0x19,0x00,0x00,0x00))}
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetVideoParameters ([int]$Brightness,[int]$Color,[int]$Contrast,[int]$Sharpness,[int]$Tint,[int]$BlackLevel,[string]$Gamma) {
        switch ($Gamma) {
            'Native'    { return $This.Command(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,0x01)}
            'SGamma'    { return $This.Command(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,0x02)}
            '2.2'       { return $This.Command(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,0x03)}
            '2.4'       { return $This.Command(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,0x04)}
            'DImage'    { return $This.Command(0x32,$Brightness,$Color,$Contrast,$Sharpness,$Tint,$BlackLevel,0x05)}
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetColorTemperature ([string]$State) {
        switch ($State) {
            'User1'     { return $This.Command((0x34,0x00))}
            'Native'    { return $This.Command((0x34,0x01))}
            '11000K'    { return $This.Command((0x34,0x02))}
            '10000K'    { return $This.Command((0x34,0x03))}
            '9300K'     { return $This.Command((0x34,0x04))}
            '7500K'     { return $This.Command((0x34,0x05))}
            '6500K'     { return $This.Command((0x34,0x06))}
            '5770K'     { return $This.Command((0x34,0x07))}
            '5500K'     { return $This.Command((0x34,0x08))}
            '5000K'     { return $This.Command((0x34,0x09))}
            '4000K'     { return $This.Command((0x34,0x0A))}
            '3400K'     { return $This.Command((0x34,0x0B))}
            '3550K'     { return $This.Command((0x34,0x0C))}
            '3000K'     { return $This.Command((0x34,0x0D))}
            '2800K'     { return $This.Command((0x34,0x0E))}
            '2600K'     { return $This.Command((0x34,0x0F))}
            '1850K'     { return $This.Command((0x34,0x10))}
            'User2'     { return $This.Command((0x34,0x12))}
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetColorParameters ([int]$Red,[int]$Green,[int]$Blue,[int]$RedOffset,[int]$GreenOffset,[int]$BlueOffset) {
        return $This.Command(0x36,$Red,$Green,$Blue,$RedOffset,$GreenOffset,$BlueOffset)
    }

    [PSCustomObject] SetPictureFormat ([string]$State) {
        switch ($State) {
            '4:3'       { return $This.Command((0x3A,0x00))}
            'Custom'    { return $This.Command((0x3A,0x01))}
            '1:1'       { return $This.Command((0x3A,0x02))}
            'Full'      { return $This.Command((0x3A,0x03))}
            '21:9'      { return $This.Command((0x3A,0x04))}
            'Dynamic'   { return $This.Command((0x3A,0x05))}
            '16:9'      { return $This.Command((0x3A,0x06))}
        }
        return [PSCustomObject]@{ Response = 'Error' }
    }

    [PSCustomObject] SetVolume ([int]$Volume,[int]$AudioOutVolume) {
        return $This.Command((0x44,$Volume,$AudioOutVolume))
    }

    [PSCustomObject] GetPower () {
        return $This.Command((0x19))
    }

    [PSCustomObject] GetIRLockStatus () {
        return $This.Command((0x1D))
    }

    [PSCustomObject] GetKeypadLockStatus () {
        return $This.Command((0x1B))
    }

    [PSCustomObject] GetReturnPowerState () {
        return $This.Command((0xA4))
    }

    [PSCustomObject] GetSource () {
        return $This.Command((0xAD))
    }

    [PSCustomObject] GetVideoParameters () {
        return $This.Command((0x33))
    }

    [PSCustomObject] GetColorTemperature () {
        return $This.Command((0x35))
    }

    [PSCustomObject] GetColorParameters () {
        return $This.Command((0x37))
    }

    [PSCustomObject] GetPictureFormat () {
        return $This.Command((0x3B))
    }

    [PSCustomObject] GetVolume () {
        return $This.Command((0x45))
    }

    [PSCustomObject] GetAudioParameters () {
        return $This.Command((0x43))
    }

    [PSCustomObject] GetOperatingHours () {
        return $This.Command((0x0F))
    }

    hidden [PSCustomObject] Command ([byte[]]$Message) {

        [byte[]]$SendBuffer = 0xA6,$This.MonitorID,0x00,0x00,0x00
        [byte[]]$ReadBuffer = @()

        $SendBuffer += [byte]$($Message.Length + 2)
        $SendBuffer += [byte]0x01
        $SendBuffer += $Message

        [byte]$XOR = 0x00
        foreach ($i in $SendBuffer) {
            [byte]$XOR = $XOR -bxor $i
        }
        
        $SendBuffer += $XOR

        $This.Device.Write($SendBuffer,0,$SendBuffer.Count)
        
        foreach ($i in 0..4)                { $ReadBuffer += $This.Device.ReadByte() }
        foreach ($i in 1..$ReadBuffer[4])   { $ReadBuffer += $This.Device.ReadByte() }
        
        if ($ReadBuffer[0] -eq 0x21) {

            if ($ReadBuffer[1] -eq $This.MonitorID) {
                
                if ($ReadBuffer[2] -eq 0x00) {

                    if ($ReadBuffer[3] -eq 0x00) {

                        $XOR = 0x00                        
                        foreach ($i in $ReadBuffer[0..($ReadBuffer.count - 2)]) {
                            [char]$XOR = $XOR -bxor $i
                        }

                        if ($ReadBuffer[-1] -ne $XOR) { return [PSCustomObject]@{ Status = 'Error' ; Details = "Bad Checksum." } }

                        switch ($ReadBuffer[6]) {
                            0x00 {
                                switch ($ReadBuffer[7]){
                                    0x00 { return [PSCustomObject]@{Status = 'Completed'                } }
                                    0x01 { return [PSCustomObject]@{Status = 'Excessive Parameters'     } }
                                    0x02 { return [PSCustomObject]@{Status = 'Insufficient Parameters'  } }
                                    0x03 { return [PSCustomObject]@{Status = 'Cancelled'                } }
                                    0x04 { return [PSCustomObject]@{Status = 'Error'                    } }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Communication Control: Property out of bounds." }
                            }
                            0x19 {
                                switch ($ReadBuffer[7]){
                                    0x01 { return [PSCustomObject]@{Power_State = 'Off' } }
                                    0x02 { return [PSCustomObject]@{Power_State = 'On'  } }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Power State: Property out of bounds." }
                            }
                            0x1D {
                                switch ($ReadBuffer[7]) {
                                    0x01 { return [PSCustomObject]@{IR_Remote_Lock = 'Unlocked'                 } }
                                    0x02 { return [PSCustomObject]@{IR_Remote_Lock = 'Locked'                   } }
                                    0x03 { return [PSCustomObject]@{IR_Remote_Lock = 'Power Unlocked'           } }
                                    0x04 { return [PSCustomObject]@{IR_Remote_Lock = 'Volume Unlocked'          } }
                                    0x05 { return [PSCustomObject]@{IR_Remote_Lock = 'Primary'                  } }
                                    0x06 { return [PSCustomObject]@{IR_Remote_Lock = 'Secondary'                } }
                                    0x07 { return [PSCustomObject]@{IR_Remote_Lock = 'Power & Volume Unlocked'  } }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "IR Lock Status: Property out of bounds." }
                            }
                            0x1B {
                                switch ($ReadBuffer[7]) {
                                    0x01 { return [PSCustomObject]@{Keypad_Lock = 'Unlocked'                } }
                                    0x02 { return [PSCustomObject]@{Keypad_Lock = 'Locked'                  } }
                                    0x03 { return [PSCustomObject]@{Keypad_Lock = 'Power Unlocked'          } }
                                    0x04 { return [PSCustomObject]@{Keypad_Lock = 'Volume Unlocked'         } }
                                    0x07 { return [PSCustomObject]@{Keypad_Lock = 'Power & Volume Unlocked' } }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Keypad Lock Status: Property out of bounds." }
                            }
                            0xA4 {
                                switch ($ReadBuffer[7]){
                                    0x01 { return [PSCustomObject]@{Return_Power_State = 'Off'          } }
                                    0x02 { return [PSCustomObject]@{Return_Power_State = 'On'           } }
                                    0x03 { return [PSCustomObject]@{Return_Power_State = 'Last Status'  } }
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Return Power Status: Property out of bounds." }
                            }
                            0xAD {
                                switch ($ReadBuffer[7]){
                                    0x01 { return [PSCustomObject]@{Current_Source = 'VIDEO'            }}
                                    0x02 { return [PSCustomObject]@{Current_Source = 'S-VIDEO'          }}
                                    0x03 { return [PSCustomObject]@{Current_Source = 'COMPONENT'        }}
                                    0x04 { return [PSCustomObject]@{Current_Source = 'CVI2'             }}
                                    0x05 { return [PSCustomObject]@{Current_Source = 'VGA'              }}
                                    0x06 { return [PSCustomObject]@{Current_Source = 'HDMI2'            }}
                                    0x07 { return [PSCustomObject]@{Current_Source = 'DisplayPort2'     }}
                                    0x08 { return [PSCustomObject]@{Current_Source = 'USB2'             }}
                                    0x09 { return [PSCustomObject]@{Current_Source = 'Card DVI-D'       }}
                                    0x0A { return [PSCustomObject]@{Current_Source = 'DisplayPort1'     }}
                                    0x0B { return [PSCustomObject]@{Current_Source = 'Card OPS'         }}
                                    0x0C { return [PSCustomObject]@{Current_Source = 'USB1'             }}
                                    0x0D { return [PSCustomObject]@{Current_Source = 'HDMI'             }}
                                    0x0E { return [PSCustomObject]@{Current_Source = 'DVI-D'            }}
                                    0x0F { return [PSCustomObject]@{Current_Source = 'HDMI3'            }}
                                    0x10 { return [PSCustomObject]@{Current_Source = 'BROWSER'          }}
                                    0x11 { return [PSCustomObject]@{Current_Source = 'SMARTCMS'         }}
                                    0x12 { return [PSCustomObject]@{Current_Source = 'DMS'              }}
                                    0x13 { return [PSCustomObject]@{Current_Source = 'INTERNAL STORAGE' }}
                                    0x14 { return [PSCustomObject]@{Current_Source = 'Reserved1'        }}
                                    0x15 { return [PSCustomObject]@{Current_Source = 'Reserved2'        }}
                                    0x16 { return [PSCustomObject]@{Current_Source = 'Media Player'     }}
                                    0x17 { return [PSCustomObject]@{Current_Source = 'PDF Player'       }}
                                    0x18 { return [PSCustomObject]@{Current_Source = 'Custom'           }}
                                    0x19 { return [PSCustomObject]@{Current_Source = 'HDMI4'            }}
                                }
                                return [PSCustomObject]@{ Status = 'Error' ; Details = "Current Source: Property out of bounds." }
                            }
                            0x33 {
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
                                                                Gamma       = switch ( $ReadBuffer[13] ) {
                                                                    0x01 {"Native"  }
                                                                    0x02 {"S-Gamma" }
                                                                    0x03 {"2.2"     }
                                                                    0x04 {"2.4"     }
                                                                    0x05 {"D-Image" }
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
                                switch ($ReadBuffer[7]){
                                    0x00 { return [PSCustomObject]@{Color_Temperature = 'User1' }}
                                    0x01 { return [PSCustomObject]@{Color_Temperature = 'Native'}}
                                    0x02 { return [PSCustomObject]@{Color_Temperature = '11000K'}}
                                    0x03 { return [PSCustomObject]@{Color_Temperature = '10000K'}}
                                    0x04 { return [PSCustomObject]@{Color_Temperature = '9300K' }}
                                    0x05 { return [PSCustomObject]@{Color_Temperature = '7500K' }}
                                    0x06 { return [PSCustomObject]@{Color_Temperature = '6500K' }}
                                    0x07 { return [PSCustomObject]@{Color_Temperature = '5770K' }}
                                    0x08 { return [PSCustomObject]@{Color_Temperature = '5500K' }}
                                    0x09 { return [PSCustomObject]@{Color_Temperature = '5000K' }}
                                    0x0A { return [PSCustomObject]@{Color_Temperature = '4000K' }}
                                    0x0B { return [PSCustomObject]@{Color_Temperature = '3400K' }}
                                    0x0C { return [PSCustomObject]@{Color_Temperature = '3350K' }}
                                    0x0D { return [PSCustomObject]@{Color_Temperature = '3000K' }}
                                    0x0E { return [PSCustomObject]@{Color_Temperature = '2800K' }}
                                    0x0F { return [PSCustomObject]@{Color_Temperature = '2600K' }}
                                    0x10 { return [PSCustomObject]@{Color_Temperature = '1850K' }}
                                    0x12 { return [PSCustomObject]@{Color_Temperature = 'User2' }}
                                }
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
                                switch ($ReadBuffer[7]){
                                    0x00 { return [PSCustomObject]@{Picture_Format = '4:3'      } }
                                    0x01 { return [PSCustomObject]@{Picture_Format = 'Custom'   } }
                                    0x02 { return [PSCustomObject]@{Picture_Format = '1:1'      } }
                                    0x03 { return [PSCustomObject]@{Picture_Format = 'Stretch'  } }
                                    0x04 { return [PSCustomObject]@{Picture_Format = '21:9'     } }
                                    0x05 { return [PSCustomObject]@{Picture_Format = 'Dynamic'  } }
                                    0x06 { return [PSCustomObject]@{Picture_Format = '16:9'     } }
                                }
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
}