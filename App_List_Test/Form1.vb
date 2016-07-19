
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSoftwareList()
    End Sub

    ' Property              Description
    ' Caption               Short description of the object.
    ' Description           Object description.
    ' IdentifyingNumber     Product identification, such As a serial number On software.
    ' InstallLocation       Location of the installed product.
    ' InstallState          Installed state Of the product. Values include:
    '                       -6 -Bad configuration
    '                       -2 - Invalid argument
    '                       -1 - Unknown package
    '                        1 - Advertised
    '                        2 - Absent
    '                        6 - Installed
    ' Name                  Commonly used product name.
    ' PackageCache          Location of the locally cached package for this product.
    ' SKUNumber             Product SKU(stock - keeping unit) information.
    ' Vendor                Name of the product's supplier.
    ' Version               Product version information.

    Private Sub LoadSoftwareList()
        ListBox1.Items.Clear()
        Dim moReturn As Management.ManagementObjectCollection
        Dim moSearch As Management.ManagementObjectSearcher
        Dim mo As Management.ManagementObject

        moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Product")

        moReturn = moSearch.Get
        For Each mo In moReturn
            On Error Resume Next ' Some stupid error making one of the items around 199 show as Null, Poor Fix
            ListBox1.Items.Add(mo("Name").ToString)
        Next
        ListBox1.Sorted = True
    End Sub
End Class
