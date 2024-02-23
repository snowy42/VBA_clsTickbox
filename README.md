# Custom Tickbox VBA Project

This project provides a custom tickbox solution for Microsoft Excel using VBA. The project consists of a class module (`clsTickBox`) and a module (`modTickBox`) for easy integration into Excel workbooks.

## Features

- **Custom Tickbox Class**: The `clsTickBox` class module allows you to create customizable tickboxes within Excel cells.
- **Easy Integration**: The `modTickBox` module provides simple subroutines to handle the click event and insert tickboxes in selected cells.

## Getting Started

### Prerequisites

- Microsoft Excel

### Installation

1. Download the project files from the GitHub repository.
2. Open your Excel workbook.
3. Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
4. Import both the `clsTickBox` class module and the `modTickBox` module into your VBA project.

## Usage

### Using the modTickBox

1. Insert a tickbox in the selected cell:
     By default, this is keybound to Ctrl+R for easy use.  This can be changed in the macro menu from the developer tab.
```vba
Public Sub Tickbox_Insert()
      Set Tickbox = New clsTickBox
 
      Dim tickShp As Shape: Set tickShp = Tickbox.CreateDefaultShape(True)
      Dim untickShp As Shape: Set untickShp = Tickbox.CreateDefaultShape(False)

      Dim targetCell As Range: Set targetCell = Selection
      Tickbox.Create tickShp, untickShp, targetCell, "Tickbox_Click"
End Sub
```
2. Handle the click event
```vba
Public Sub Tickbox_Click()
     ' Handles the click event for the tickbox

     ' Check if the tick object is already instantiated
     ' this will likely only call when the workbook is reopened
     If Tickbox Is Nothing Then Set Tickbox = New clsTickBox

     ' Call the Click method of the tick object passing the caller information
     Tickbox.Click Application.caller
End Sub
```

### Customization

The Tickbox.CreateDefaultShape function will create a default tickbox look.  
If you want to use a custom shape or style for the tickboxes, then simply create the shape or import the image into excel and assign the ticked and unticked version as "tickShp" and "untickShp" in the Insert routine.  This will accept shape groups as well as individual shapes.
```vba
Public Sub Tickbox_Insert()
      Set Tickbox = New clsTickBox
 
      Dim tickShp As Shape: Set tickShp = Sheet1.Shapes("my_Ticked_Shape") ' adjust as needed
      Dim untickShp As Shape: Set untickShp = Sheet1.Shapes("my_Unticked_Shape") ' adjust as needed

      Dim targetCell As Range: Set targetCell = Selection
      Tickbox.Create tickShp, untickShp, targetCell, "Tickbox_Click"
End Sub
```


## Contributing

Contributions are welcome!  IF you have suggestions or improvements, feel free to open an issue or submit a pull request.
