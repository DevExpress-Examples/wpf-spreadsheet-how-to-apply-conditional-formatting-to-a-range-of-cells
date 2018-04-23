Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports DevExpress.Spreadsheet
Imports System.Drawing
Imports DevExpress.Utils

Namespace ConditionalFormatting_WPF_Examples_VB
    Friend Class ConditionalFormatting
#Region "Actions"
        Public Shared AddAverageConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddAverageConditionalFormatting
        Public Shared AddRangeConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddRangeConditionalFormatting
        Public Shared AddRankConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddRankConditionalFormatting
        Public Shared AddTextConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddTextConditionalFormatting
        Public Shared AddSpecialConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddSpecialConditionalFormatting
        Public Shared AddTimePeriodConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddTimePeriodConditionalFormatting
        Public Shared AddExpressionConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddExpressionConditionalFormatting
        Public Shared AddFormulaExpressionConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddFormulaExpressionConditionalFormatting
        Public Shared AddColorScale2ConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddColorScale2ConditionalFormatting
        Public Shared AddColorScale3ConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddColorScale3ConditionalFormatting
        Public Shared AddDataBarConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddDataBarConditionalFormatting
        Public Shared AddIconSetConditionalFormattingAction As Action(Of IWorkbook) = AddressOf AddIconSetConditionalFormatting
#End Region

        Private Shared Sub AddAverageConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '				#Region "#AverageConditionalFormatting"
                Dim conditionalFormattings As ConditionalFormattingCollection = worksheet.ConditionalFormattings
                ' Create the rule highlighting values that are above the average in cells C2 through C15.  
                Dim cfRule1 As AverageConditionalFormatting = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range("$C$2:$C$15"), ConditionalFormattingAverageCondition.AboveOrEqual)
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to yellow.
                cfRule1.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HFA, &HF7, &HAA)
                ' Set the font color to red.
                cfRule1.Formatting.Font.Color = Color.Red
                ' Create the rule highlighting values that are one standard deviation below the mean in cells D2 through D15.
                Dim cfRule2 As AverageConditionalFormatting = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range("$D$2:$D$15"), ConditionalFormattingAverageCondition.BelowOrEqual, 1)
                ' Specify formatting options to be applied to cells if the conditions is true.
                ' Set the background color to light-green.
                cfRule2.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &H9F, &HFB, &H69)
                ' Set the font color to blue-violet.
                cfRule2.Formatting.Font.Color = Color.BlueViolet
                '				#End Region ' #AverageConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Determine cost values that are above the average in the first quarter and one standard deviation below the mean in the second quarter."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub AddRangeConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#RangeConditionalFormatting"
                ' Create the rule to identify values below 7 and above 19 in cells F2 through F15.  
                Dim cfRule As RangeConditionalFormatting = worksheet.ConditionalFormattings.AddRangeConditionalFormatting(worksheet.Range("$F$2:$F$15"), ConditionalFormattingRangeCondition.Outside, "7", "19")
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to yellow.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HFA, &HF7, &HAA)
                ' Set the font color to red.
                cfRule.Formatting.Font.Color = Color.Red
                '			#End Region ' #RangeConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Identify book prices that are below $7 and above $19."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub AddRankConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#RankConditionalFormatting"
                ' Create the rule to identify top three values in cells F2 through F15.
                Dim cfRule As RankConditionalFormatting = worksheet.ConditionalFormattings.AddRankConditionalFormatting(worksheet.Range("$F$2:$F$15"), ConditionalFormattingRankCondition.TopByRank, 3)
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to dark orchid.
                cfRule.Formatting.Fill.BackgroundColor = Color.DarkOrchid
                ' Set the outline borders.
                cfRule.Formatting.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin)
                ' Set the font color to white.
                cfRule.Formatting.Font.Color = Color.White
                '			#End Region ' #RankConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Identify the top three price values."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub AddTextConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#TextConditionalFormatting"
                ' Create the rule to highlight values with the given text string in cells A2 through A15.
                Dim cfRule As TextConditionalFormatting = worksheet.ConditionalFormattings.AddTextConditionalFormatting(worksheet.Range("$A$2:$A$15"), ConditionalFormattingTextCondition.Contains, "Bradbury")
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to pink.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HE1, &H95, &HC2)
                '			#End Region ' #TextConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Quickly find books written by Ray Bradbury."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub


        Private Shared Sub AddSpecialConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#SpecialConditionalFormatting"
                ' Create the rule to identify unique values in cells A2 through A15. 
                Dim cfRule As SpecialConditionalFormatting = worksheet.ConditionalFormattings.AddSpecialConditionalFormatting(worksheet.Range("$A$2:$A$15"), ConditionalFormattingSpecialCondition.ContainUniqueValue)
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to yellow.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HFA, &HF7, &HAA)
                '			#End Region ' #SpecialConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "In a list of authors quickly identify unique values."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub


        Private Shared Sub AddTimePeriodConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfTasks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#TimePeriodConditionalFormatting"
                ' Create the rule to highlight today's dates in cells B2 through B6.
                Dim cfRule As TimePeriodConditionalFormatting = worksheet.ConditionalFormattings.AddTimePeriodConditionalFormatting(worksheet.Range("$B$2:$B$6"), ConditionalFormattingTimePeriod.Today)
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to pink.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HF2, &HAE, &HE3)
                '			#End Region ' #TimePeriodConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A8:B9")
                ruleExplanation.Value = "Determine the today's task in the TO DO list."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub


        Private Shared Sub AddExpressionConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#ExpressionConditionalFormatting"
                ' Create the rule to identify values that are above the average in cells F2 through F15.
                Dim cfRule As ExpressionConditionalFormatting = worksheet.ConditionalFormattings.AddExpressionConditionalFormatting(worksheet.Range("$F$2:$F$15"), ConditionalFormattingExpressionCondition.GreaterThan, "=AVERAGE($F$2:$F$15)")
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to yellow.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HFA, &HF7, &HAA)
                ' Set the font color to red.
                cfRule.Formatting.Font.Color = Color.Red
                '			#End Region ' #ExpressionConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Identify book prices that are greater than the average price."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub


        Private Shared Sub AddFormulaExpressionConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#FormulaExpressionConditionalFormatting"
                ' Create the rule to shade alternate rows without applying a new style.
                Dim cfRule As FormulaExpressionConditionalFormatting = worksheet.ConditionalFormattings.AddFormulaExpressionConditionalFormatting(worksheet.Range("$A$2:$G$15"), "=MOD(ROW(),2)=1")
                ' Specify formatting options to be applied to cells if the condition is true.
                ' Set the background color to light blue.
                cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, &HBC, &HDA, &HF7)
                '			#End Region ' #FormulaExpressionConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Shade alternate rows in light blue without applying a new style."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub AddColorScale2ConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#ColorScale2ConditionalFormatting"
                Dim conditionalFormattings As ConditionalFormattingCollection = worksheet.ConditionalFormattings
                ' Set the minimum threshold to the lowest value in the range of cells.
                Dim minPoint As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
                ' Set the maximum threshold to the highest value in the range of cells.
                Dim maxPoint As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
                ' Create the two-color scale rule to differentiate low and high values in cells C2 through D15. Blue represents the lower values and yellow represents the higher values. 
                Dim cfRule As ColorScale2ConditionalFormatting = conditionalFormattings.AddColorScale2ConditionalFormatting(worksheet.Range("$C$2:$D$15"), minPoint, Color.FromArgb(255, &H9D, &HE9, &HFA), maxPoint, Color.FromArgb(255, &HFF, &HF6, &HA9))
                '			#End Region ' #ColorScale2ConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Examine cost distribution using a gradation of two colors. Blue represents the lower values and yellow represents the higher values."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub


        Private Shared Sub AddColorScale3ConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#ColorScale3ConditionalFormatting"
                Dim conditionalFormattings As ConditionalFormattingCollection = worksheet.ConditionalFormattings
                ' Set the minimum threshold to the lowest value in the range of cells using the MIN() formula.
                Dim minPoint As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Formula, "=MIN($C$2:$D$15)")
                ' Set the midpoint threshold to the 50th percentile.
                Dim midPoint As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percentile, "50")
                ' Set the maximum threshold to the highest value in the range of cells using the MAX() formula.
                Dim maxPoint As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Number, "=MAX($C$2:$D$15)")
                ' Create the three-color scale rule to determine how values in cells C2 through D15 vary. Red represents the lower values, yellow represents the medium values and sky blue represents the higher values.
                Dim cfRule As ColorScale3ConditionalFormatting = conditionalFormattings.AddColorScale3ConditionalFormatting(worksheet.Range("$C$2:$D$15"), minPoint, Color.Red, midPoint, Color.Yellow, maxPoint, Color.SkyBlue)
                '			#End Region ' #ColorScale3ConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Examine cost distribution using a gradation of three colors. Red represents the lower values, yellow represents the medium values and sky blue represents the higher values."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub AddDataBarConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#DataBarConditionalFormatting"
                Dim conditionalFormattings As ConditionalFormattingCollection = worksheet.ConditionalFormattings
                ' Set the value corresponding to the shortest bar to the lowest value.
                Dim lowBound1 As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
                ' Set the value corresponding to the longest bar to the highest value.
                Dim highBound1 As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
                ' Create the rule to compare values in cells E2 through E15 using data bars. 
                Dim cfRule1 As DataBarConditionalFormatting = conditionalFormattings.AddDataBarConditionalFormatting(worksheet.Range("$E$2:$E$15"), lowBound1, highBound1, DXColor.Green)
                ' Set the positive bar border color to green.
                cfRule1.BorderColor = DXColor.Green
                ' Set the negative bar color to red.
                cfRule1.NegativeBarColor = DXColor.Red
                ' Set the negative bar border color to red.
                cfRule1.NegativeBarBorderColor = DXColor.Red
                ' Set the axis position to display the axis in the middle of the cell.
                cfRule1.AxisPosition = ConditionalFormattingDataBarAxisPosition.Middle
                ' Set the axis color to dark blue.
                cfRule1.AxisColor = Color.DarkBlue

                ' Set the value corresponding to the shortest bar to 0 percent.
                Dim lowBound2 As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percent, "0")
                ' Set the value corresponding to the longest bar to 100 percent.
                Dim highBound2 As ConditionalFormattingValue = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percent, "100")
                ' Create the rule to compare values in cells G2 through G15 using data bars.  
                Dim cfRule2 As DataBarConditionalFormatting = conditionalFormattings.AddDataBarConditionalFormatting(worksheet.Range("$G$2:$G$15"), lowBound2, highBound2, DXColor.SkyBlue)
                ' Set the data bar border color to sky blue.
                cfRule2.BorderColor = DXColor.SkyBlue
                ' Specify the solid fill type.
                cfRule2.GradientFill = False
                ' Hide values of cells to which the rule is applied.
                cfRule2.ShowValue = False
                '			#End Region ' #DataBarConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Compare values in the ""Cost Trend"" and ""Markup"" columns using data bars."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub
        Private Shared Sub AddIconSetConditionalFormatting(ByVal workbook As IWorkbook)
            workbook.Calculate()
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets("cfBooks")
                workbook.Worksheets.ActiveWorksheet = worksheet
                '			#Region "#IconSetConditionalFormatting"
                Dim conditionalFormattings As ConditionalFormattingCollection = worksheet.ConditionalFormattings
                ' Set the first threshold to the lowest value in the range of cells using the MIN() formula.
                Dim minPoint As ConditionalFormattingIconSetValue = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Formula, "=MIN($E$2:$E$15)", ConditionalFormattingValueOperator.GreaterOrEqual)
                ' Set the second threshold to 0.
                Dim midPoint As ConditionalFormattingIconSetValue = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Number, "0", ConditionalFormattingValueOperator.GreaterOrEqual)
                ' Set the third threshold to 0.01.
                Dim maxPoint As ConditionalFormattingIconSetValue = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Number, "0.01", ConditionalFormattingValueOperator.GreaterOrEqual)
                ' Create the rule to apply a specific icon from the three arrow icon set to each cell in the range  E2:E15 based on its value.  
                Dim cfRule As IconSetConditionalFormatting = conditionalFormattings.AddIconSetConditionalFormatting(worksheet.Range("$E$2:$E$15"), IconSetType.Arrows3, New ConditionalFormattingIconSetValue() {minPoint, midPoint, maxPoint})
                ' Specify the custom icon to be displayed if the second condition is true. 
                ' To do this, set the IconSetConditionalFormatting.IsCustom property to true, which is false by default.
                cfRule.IsCustom = True
                ' Initialize the ConditionalFormattingCustomIcon object.
                Dim cfCustomIcon As New ConditionalFormattingCustomIcon()
                ' Specify the icon set where you wish to get the icon. 
                cfCustomIcon.IconSet = IconSetType.TrafficLights13
                ' Specify the index of the desired icon in the set.
                cfCustomIcon.IconIndex = 1
                ' Add the custom icon at the specified position in the initial icon set.
                cfRule.SetCustomIcon(1, cfCustomIcon)
                ' Hide values of cells to which the rule is applied.
                cfRule.ShowValue = False
                '			#End Region ' #IconSetConditionalFormatting
                ' Add an explanation to the created rule.
                Dim ruleExplanation As Range = worksheet.Range("A17:G18")
                ruleExplanation.Value = "Identify upward and downward cost trends."
            Finally
                workbook.EndUpdate()
            End Try
        End Sub
    End Class
End Namespace
