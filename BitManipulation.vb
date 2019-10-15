''' <summary>
''' Provides basic bit manipulation functionality. 
''' </summary>
Public Module BitManipulation
    '--------------------------MakeWord-------------------------------------
    ''' <summary>
    ''' Makes a word from two bytes.
    ''' </summary>
    ''' <param name="loByte">The low byte.</param>
    ''' <param name="hiByte">The high byte.</param>
    ''' <returns>The combined bytes as a word.</returns>
    ''' <remarks>Works for all the possible combinations for the low and high order byte from <see cref="System.Byte.MinValue">0</see> through <see cref="System.Byte.MaxValue">255</see>.</remarks>
    Public Function MakeWord(ByVal loByte As System.Byte, _
                             ByVal hiByte As System.Byte) _
                             As System.Int16
        '&H80           =  128      = 00000000 00000000 00000000 10000000 
        '&H100          =  256      = 00000000 00000000 00000001 00000000
        '&HFFFF0000     = -65536    = 11111111 11111111 00000000 00000000 
        If CBool(hiByte And &H80) Then
            Return CType(((hiByte * &H100&) + loByte) Or &HFFFF0000, System.Int16)
        Else
            Return CType((hiByte * &H100) + loByte, System.Int16)
        End If
    End Function

    '--------------------------HiByte-------------------------------------
    ''' <summary>
    ''' Gets the high byte from a given word.
    ''' </summary>
    ''' <param name="word">The word.</param>
    ''' <returns>The high byte from the word.</returns>
    ''' <remarks>Works for all the possible values for the word from <see cref="Int16.MinValue">-32,768</see> through <see cref="Int16.MaxValue">32,767</see>.</remarks>
    Public Function HiByte(ByVal word As System.Int16) As System.Byte
        '&HFF00&        =  65280     = 00000000 00000000 11111111 00000000 
        '&H100          =  256       = 00000000 00000000 00000001 00000000
        'Integer Division Here
        Return CType((word And &HFF00&) \ &H100, System.Byte)
    End Function

    '--------------------------LoByte-------------------------------------
    ''' <summary>
    ''' Gets the low byte from a given word.
    ''' </summary>
    ''' <param name="word">The word.</param>
    ''' <returns>The low byte from the word.</returns>
    ''' <remarks>Works for all the possible values for the word from <see cref="Int16.MinValue">-32,768</see> through <see cref="Int16.MaxValue">32,767</see>.</remarks>
    Function LoByte(ByVal word As System.Int16) As System.Byte
        '&HFF           =  255       = 00000000 00000000 00000000 11111111 
        Return CType(word And &HFF, System.Byte)
    End Function

    '--------------------------HiWord-------------------------------------
    ''' <summary>
    ''' Gets the high word from a given double word.
    ''' </summary>
    ''' <param name="dWord">The double word.</param>
    ''' <returns>The high word from the double word.</returns>
    ''' <remarks>Works for all the possible values for the dWord from <see cref="Int32.MinValue">-2,147,483,648</see> through <see cref="Int32.MaxValue">2,147,483,647</see>.</remarks>
    Public Function HiWord(ByVal dWord As System.Int32) As System.Int16
        '&HFFFF0000     = -65536    = 11111111 11111111 00000000 00000000 
        '&H10000        =  65536    = 00000000 00000001 00000000 00000000
        'Integer Division Here.
        Return CType((dWord And &HFFFF0000) \ &H10000, System.Int16)
    End Function

    '--------------------------LoWord------------------------------------
    ''' <summary>
    ''' Gets the low word from a given double word.
    ''' </summary>
    ''' <param name="dWord">The double word.</param>
    ''' <returns>The low word from the double word.</returns>
    ''' <remarks>Works for all the possible values for the dWord from <see cref="Int32.MinValue">-2,147,483,648</see> through <see cref="Int32.MaxValue">2,147,483,647</see>.</remarks>
    Public Function LoWord(ByVal dWord As System.Int32) As System.Int16
        '&H8000         =  32768   = 00000000 00000000 10000000 00000000 
        'Test if the dWord parameter is a positive or negative integer value.
        If CType(dWord And &H8000&, System.Boolean) Then
            'The dWord is positive.
            '&HFFFF0000     = -65536    = 11111111 11111111 00000000 00000000 
            Return CType(dWord Or &HFFFF0000, System.Int16)
        Else
            'The dWord is negative.
            '&HFFFF&        =  65535    = 00000000 00000000 11111111 11111111
            Return CType(dWord And &HFFFF&, System.Int16)
        End If
    End Function

    '--------------------------MakeDWord/MakeLong------------------------------------
    ''' <summary>
    ''' Makes a double word from two words.
    ''' </summary>
    ''' <param name="loWord">The low byte.</param>
    ''' <param name="hiWord">The high byte.</param>
    ''' <returns>The combined bytes as a word.</returns>
    ''' <remarks>Works for all the possible combinations for the low and high order word from <see cref="Int16.MinValue">-32768</see> through <see cref="Int16.MaxValue">32768</see>.</remarks>
    Public Function MakeDWord(ByVal loWord As System.Int16, _
                              ByVal hiWord As System.Int16) _
                              As System.Int32
        '&H10000        =  65536        = 00000000 00000001 00000000 00000000
        '&HFFFF&        =  65535        = 00000000 00000000 11111111 11111111
        Return CType((hiWord * &H10000) Or (loWord And &HFFFF&), System.Int32)
    End Function

    '**********************************************************************
End Module

