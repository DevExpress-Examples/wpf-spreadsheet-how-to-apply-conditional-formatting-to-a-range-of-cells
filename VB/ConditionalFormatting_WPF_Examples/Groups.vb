Imports System.Collections.Generic

Namespace ConditionalFormatting_WPF_Examples

    Public Partial Class Groups
        Inherits List(Of Group)

        Public Shared Function InitData() As Groups
            Dim examples As Groups = New Groups()
#Region "GroupNodes"
            examples.Add(New Group("Conditional Formatting Examples"))
#End Region
#Region "ExampleNodes"
            ' Add nodes to the "Conditional Formatting" group of examples.
            examples(0).Items.Add(New SpreadsheetExample("Format values that are above or below average", ConditionalFormatting.AddAverageConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells that are not between two specified values", ConditionalFormatting.AddRangeConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format top ranked values", ConditionalFormatting.AddRankConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells that contain the given text", ConditionalFormatting.AddTextConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format only unique values", ConditionalFormatting.AddSpecialConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format today's date", ConditionalFormatting.AddTimePeriodConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format values that are greater than a specified value", ConditionalFormatting.AddExpressionConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Use a formula to determine which cells to format", ConditionalFormatting.AddFormulaExpressionConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells by using a two-color scale", ConditionalFormatting.AddColorScale2ConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells by using a three-color scale", ConditionalFormatting.AddColorScale3ConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells by using data bars", ConditionalFormatting.AddDataBarConditionalFormattingAction))
            examples(0).Items.Add(New SpreadsheetExample("Format cells by using a custom icon set", ConditionalFormatting.AddIconSetConditionalFormattingAction))
            Return examples
#End Region
        End Function
    End Class
End Namespace
