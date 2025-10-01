Attribute VB_Name = "Module11"
Sub InventoryAddingFormat()

    ' Tracks index, used for both column and row
    Dim i As Long
        i = 1
    
    Const MARKUP_PERCENT As Double = 0.35 ' Used for easy editing of the markup in future
    
    Const SALE_PERCENT As Double = 0.15 ' Used for easy editing of the markup in future, any changes to this number could put the price below the MAP price
    
    
    ' Removing unneeded fields, any field in the list will be be deleted
    Dim trashList As Variant
    trashList = Array("Pickup Price with Prepaid Shipping Label (US$)", "Promotion Flag", "Sale Price (US$)", "Sale Price for Pickup (US$)", "Promotion Start Date PST", "Promotion End Date PST")
    
    Do While ActiveSheet.Cells(1, i).Value <> "" ' Do while there are still headers, will check cells A1, B1, C1... until it finds an empty one
        For j = LBound(trashList) To UBound(trashList) ' For every string in the trashlist array
            If ActiveSheet.Cells(1, i).Value = trashList(j) Then ' If a match is found
                ActiveSheet.Columns(i).Delete ' Delete the entire column
                i = i - 1 ' Deleting a column shifts all the later columns 1 to the left so we have to shift the index as well
                Exit For
            End If
        Next j
        i = i + 1
    Loop
    
        
    ' Tag important columns
        ' Will store the index of the following columns: "Dropshipping Price (US$)", "Estimate Shipping Cost (US$)", "MAP (US$)", and the first empty column to use in pricing, pricing, verifying pricing, and knowing bounds respectivly
        Dim PriceIndex As Long
        Dim ShippingPriceIndex As Long
        Dim MAPIndex As Long
        Dim EndIndex As Long
    
        EndIndex = i ' i is currently the first empty column
        i = 1 ' reseting i
        
        Do While ActiveSheet.Cells(1, i).Value <> "" ' Do while the cell at coordinates (1, i) (aka: A1, B1, C1 ...) is not empty
            If ActiveSheet.Cells(1, i).Value = "Dropshipping Price (US$)" Then
                PriceIndex = i
            End If
            If ActiveSheet.Cells(1, i).Value = "Estimate Shipping Cost (US$)" Then
                ShippingPriceIndex = i
            End If
            If ActiveSheet.Cells(1, i).Value = "MAP (US$)" Then
                MAPIndex = i
            End If
            i = i + 1
        Loop
    

    ' Creating new price fields
        ' Create new field titled "Price(35%)", the number may be different depending on if the markup price is changed
        ' Fills field with the math formula: (Price + 35% markup) + Shipping
        
    ActiveSheet.Cells(1, EndIndex).Value = "Price(" & CInt(MARKUP_PERCENT * 100) & "%)" ' Names the field
    
    i = 2 ' i is reset and is now used to track the row index
    
    Do While ActiveSheet.Cells(i, 1).Value <> "" ' Do while there are still rows left to price
    
        ' Some cells don't have numerical values, when this happens, they default to 0 which is fine because there is only a non-numerical number when it should be 0
        Dim price As Double
        Dim shipping As Double
        
        price = CDbl(ActiveSheet.Cells(i, PriceIndex).Value) ' Prices are always numeric, no need to check
        
        If IsNumeric(ActiveSheet.Cells(i, ShippingPriceIndex).Value) Then ' Checks if the shipping price is numeric
            shipping = CDbl(ActiveSheet.Cells(i, ShippingPriceIndex).Value)
        Else
            shipping = 0 ' Default to 0 if the value is not a number e.g. "N/A" or ""
        End If

        ActiveSheet.Cells(i, EndIndex).Value = Round((price + (price * MARKUP_PERCENT) + shipping), 2)
        i = i + 1
    Loop
    
    
    ' Verifying prices
        ' Making sure the new prices are higher than those in the "MAP (US$) field
        
    i = 2 ' Reseting i
        
    Do While ActiveSheet.Cells(i, 1).Value <> "" ' Do while there are still rows left to verify
        Dim mapVal As Double
        Dim priceVal As Double
        
        If IsNumeric(ActiveSheet.Cells(i, MAPIndex).Value) Then
            mapVal = CDbl(ActiveSheet.Cells(i, MAPIndex).Value)
        Else
            mapVal = 0 ' Default to 0 if the value is not a number e.g. "N/A" or ""
        End If
        
        ' Converts the price value into a double
        If IsNumeric(ActiveSheet.Cells(i, EndIndex - 1).Value) Then
            priceVal = CDbl(ActiveSheet.Cells(i, EndIndex - 1).Value)
        End If
        
        If priceVal - (priceVal * SALE_PERCENT) <= mapVal Then ' If the price when on sale is lower than the MAP price, we have to have it higher or we're going to get sued
        
            Dim newPrice As Double ' Price the item will be set to
            newPrice = mapVal / (1 - SALE_PERCENT) ' The formula to make sure even a sale of SALE_PERCENT doesn't make the price go under the MAP price
            
            ActiveSheet.Cells(i, EndIndex).Value = Round(newPrice + 1, 2) ' Adding one more dollar just in case
        End If
        i = i + 1
    Loop
    
    EndIndex = EndIndex + 1 ' Setting the end index to the next open line
    
    ' Second Verification
        ' Flags the row next to the price as "ERROR" if the number is wrong (price when on sale <= MAP price)
        
    ActiveSheet.Cells(1, EndIndex).Value = "MAP Errors"
    
    i = 2 ' Reseting i
    
    Do While ActiveSheet.Cells(i, 1).Value <> ""
        Dim mapVal2 As Double
        Dim priceval2 As Double
        
        If IsNumeric(ActiveSheet.Cells(i, MAPIndex).Value) Then
            mapVal2 = CDbl(ActiveSheet.Cells(i, MAPIndex).Value)
        Else
            mapVal2 = 0 ' Default to 0 if the value is not a number (e.g., "N/A", "")
        End If
        
        ' Converts the price value into a double
        If IsNumeric(ActiveSheet.Cells(i, EndIndex - 1).Value) Then
            priceval2 = CDbl(ActiveSheet.Cells(i, EndIndex - 1).Value)
        End If
        
        If priceval2 - (priceval2 * SALE_PERCENT) <= mapVal2 Then
            ActiveSheet.Cells(i, EndIndex).Value = "ERROR"
        End If
        
        i = i + 1
    Loop
    
    EndIndex = EndIndex + 1
    
End Sub
