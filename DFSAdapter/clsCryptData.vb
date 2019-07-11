Imports System.Security.Cryptography
Public Class clsCryptData
    Const StrKey As String = "Citadel264"
    Private Shared Function CreateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider
        ' Hash the key.
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)
        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    Public Shared Function EncodingData(ByVal strCodingText As String)
        Try
            If strCodingText.Length = 0 Then
                Return strCodingText
            End If
            Dim TripleDes As New TripleDESCryptoServiceProvider
            TripleDes.Key = CreateHash(StrKey, TripleDes.KeySize \ 8)
            TripleDes.IV = CreateHash("", TripleDes.BlockSize \ 8)
            ' Convert the plaintext string to a byte array.
            Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(strCodingText)
            ' Create the stream.
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream.
            Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()
            ' Convert the encrypted stream to a printable string.
            Return Convert.ToBase64String(ms.ToArray)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function DecodingData(ByVal strCodingText As String)
        Try
            If strCodingText.Length = 0 Then
                Return strCodingText
            End If
            Dim TripleDes As New TripleDESCryptoServiceProvider
            TripleDes.Key = CreateHash(StrKey, TripleDes.KeySize \ 8)
            TripleDes.IV = CreateHash("", TripleDes.BlockSize \ 8)
            ' Convert the encrypted text string to a byte array.
            Dim encryptedBytes() As Byte = Convert.FromBase64String(strCodingText)
            ' Create the stream.
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream.
            Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()
            ' Convert the plaintext stream to a string.
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        Catch ex As Exception
            Return ""
        End Try
    End Function
End Class

