Attribute VB_Name = "Populate"
Option Explicit

Public Sub populate_shoppinglist()

    Dim BreakfastArea As Range
    Set BreakfastArea = wsPlan.Range("BreakfastArea")
    
    Dim SnackAreaAM As Range
    Set SnackAreaAM = wsPlan.Range("SnacksAreaAM")
    
    Dim LunchArea As Range
    Set LunchArea = wsPlan.Range("LunchArea")
    
    Dim SnackAreaPM As Range
    Set SnackAreaPM = wsPlan.Range("SnacksAreaPM")
    
    Dim DinnerArea As Range
    Set DinnerArea = wsPlan.Range("DinnerArea")
    
    Dim ListArea As Range
    Set ListArea = wsPlan.Range("ListArea")
    ListArea.ClearContents
    
    Dim ShoppingArea As Range
    Set ShoppingArea = wsShopping.Columns(1)
    ShoppingArea.ClearContents
    
    Dim Ingredients() As String
    Dim ShoppingRow As Long
    Dim ShoppingLastRow As Long
    Dim ArrayItem As Long
    Dim ListColumn As Long
    Dim ListRow As Long
    
    'The ShoppingRow keeps track of the current row on the wsShopping as we compile the ingredient list
    ShoppingRow = 1

    ShoppingRow = FindIngredients(wsBreakfast, BreakfastArea, ShoppingRow)
    ShoppingRow = FindIngredients(wsSnacks, SnackAreaAM, ShoppingRow)
    ShoppingRow = FindIngredients(wsLunch, LunchArea, ShoppingRow)
    ShoppingRow = FindIngredients(wsSnacks, SnackAreaPM, ShoppingRow)
    ShoppingRow = FindIngredients(wsDinner, DinnerArea, ShoppingRow)
 
    'Many food items have the same ingredients
    wsShopping.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo

    'Essentially checking for no selections on wsPlan but checking wsShopping is easier because of the data validation
    If IsEmpty(wsShopping.Range("A1")) Then
        MsgBox ("No selections")
    Exit Sub
    End If
    
    ShoppingLastRow = wsShopping.Cells(Rows.Count, 1).End(xlUp).Row
    ReDim Ingredients(1 To ShoppingLastRow)
    
    For ArrayItem = 1 To ShoppingLastRow
        Ingredients(ArrayItem) = wsShopping.Cells(ArrayItem, 1)
    Next
    
    ShoppingRow = 1
    ListColumn = 2
    
Populate:
        On Error GoTo Finish
        For ListRow = 14 To 29
            wsPlan.Cells(ListRow, ListColumn) = Ingredients(ShoppingRow)
            ShoppingRow = ShoppingRow + 1
        Next
    
        If ShoppingRow - 1 < ShoppingLastRow Then
            ListColumn = ListColumn + 1
            GoTo Populate
        End If
    
Finish:
        wsShopping.Range("A:A").Clear
    
    End Sub


Public Function FindIngredients(ByVal IngredientSheet As Worksheet, ByVal FoodRange As Range, ByVal ShoppingRow As Long) As Long
'This subroutine takes all of the selections in an area and finds the ingredients for each selection


    Dim FoodSelection As Range
    Dim Ingredient As Range
    
    Dim ColumnNumber As Long
    Dim RowNumber As Long
    Dim ColumnShoppingRow As Long

    For Each FoodSelection In FoodRange
    
        If FoodSelection.Value <> "" Then
            Set Ingredient = IngredientSheet.Range("A:A").Find(FoodSelection.Value, LookIn:=xlValues, lookat:=xlWhole)
            If Not Ingredient Is Nothing Then
                RowNumber = Ingredient.Row
                ColumnNumber = Ingredient.End(xlToRight).Column
                    For ColumnShoppingRow = 2 To ColumnNumber
                        wsShopping.Cells(ShoppingRow, 1) = IngredientSheet.Cells(RowNumber, ColumnShoppingRow)
                        ShoppingRow = ShoppingRow + 1
                    Next ColumnShoppingRow
            End If
        End If
    Next FoodSelection
    FindIngredients = ShoppingRow
End Function
