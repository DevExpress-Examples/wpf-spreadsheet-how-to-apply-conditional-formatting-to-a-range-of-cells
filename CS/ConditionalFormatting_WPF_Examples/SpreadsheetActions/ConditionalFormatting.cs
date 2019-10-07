using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Spreadsheet;
using System.Drawing;
using DevExpress.Utils;

namespace ConditionalFormatting_WPF_Examples {
    class ConditionalFormatting {
        #region Actions
        public static Action<IWorkbook> AddAverageConditionalFormattingAction = AddAverageConditionalFormatting;
        public static Action<IWorkbook> AddRangeConditionalFormattingAction = AddRangeConditionalFormatting;
        public static Action<IWorkbook> AddRankConditionalFormattingAction = AddRankConditionalFormatting;
        public static Action<IWorkbook> AddTextConditionalFormattingAction = AddTextConditionalFormatting;
        public static Action<IWorkbook> AddSpecialConditionalFormattingAction = AddSpecialConditionalFormatting;
        public static Action<IWorkbook> AddTimePeriodConditionalFormattingAction = AddTimePeriodConditionalFormatting;
        public static Action<IWorkbook> AddExpressionConditionalFormattingAction = AddExpressionConditionalFormatting;
        public static Action<IWorkbook> AddFormulaExpressionConditionalFormattingAction = AddFormulaExpressionConditionalFormatting;
        public static Action<IWorkbook> AddColorScale2ConditionalFormattingAction = AddColorScale2ConditionalFormatting;
        public static Action<IWorkbook> AddColorScale3ConditionalFormattingAction = AddColorScale3ConditionalFormatting;
        public static Action<IWorkbook> AddDataBarConditionalFormattingAction = AddDataBarConditionalFormatting;
        public static Action<IWorkbook> AddIconSetConditionalFormattingAction = AddIconSetConditionalFormatting;
        #endregion

        static void AddAverageConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                workbook.Worksheets.ActiveWorksheet = worksheet;
                #region #AverageConditionalFormatting
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Create the rule highlighting values that are above the average in cells C2 through C15.  
                AverageConditionalFormatting cfRule1 = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range["$C$2:$C$15"], ConditionalFormattingAverageCondition.AboveOrEqual);
                // Specify formatting options to be applied to cells if the condition is true.
                // Set the background color to yellow.
                cfRule1.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xFA, 0xF7, 0xAA);
                // Set the font color to red.
                cfRule1.Formatting.Font.Color = Color.Red;
                // Create the rule highlighting values that are one standard deviation below the mean in cells D2 through D15.
                AverageConditionalFormatting cfRule2 = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range["$D$2:$D$15"], ConditionalFormattingAverageCondition.BelowOrEqual, 1);
                // Specify formatting options to be applied to cells if the conditions is true.
                // Set the background color to light-green.
                cfRule2.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0x9F, 0xFB, 0x69);
                // Set the font color to blue-violet.
                cfRule2.Formatting.Font.Color = Color.BlueViolet;
                #endregion #AverageConditionalFormatting
                // Add an explanation to the created rule.
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Determine cost values that are above the average in the first quarter and one standard deviation below the mean in the second quarter.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddRangeConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #RangeConditionalFormatting
            // Create the rule to identify values below 7 and above 19 in cells F2 through F15.  
            RangeConditionalFormatting cfRule = worksheet.ConditionalFormattings.AddRangeConditionalFormatting(worksheet.Range["$F$2:$F$15"], ConditionalFormattingRangeCondition.Outside, "7", "19");
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to yellow.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xFA, 0xF7, 0xAA);
            // Set the font color to red.
            cfRule.Formatting.Font.Color = Color.Red;
            #endregion #RangeConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "Identify book prices that are below $7 and above $19.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddRankConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #RankConditionalFormatting
            // Create the rule to identify top three values in cells F2 through F15.
            RankConditionalFormatting cfRule = worksheet.ConditionalFormattings.AddRankConditionalFormatting(worksheet.Range["$F$2:$F$15"], ConditionalFormattingRankCondition.TopByRank, 3);
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to dark orchid.
            cfRule.Formatting.Fill.BackgroundColor = Color.DarkOrchid;
            // Set the outline borders.
            cfRule.Formatting.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);
            // Set the font color to white.
            cfRule.Formatting.Font.Color = Color.White;
            #endregion #RankConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "Identify the top three price values.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddTextConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #TextConditionalFormatting
            // Create the rule to highlight values with the given text string in cells A2 through A15.
            TextConditionalFormatting cfRule = worksheet.ConditionalFormattings.AddTextConditionalFormatting(worksheet.Range["$A$2:$A$15"], ConditionalFormattingTextCondition.Contains, "Bradbury");
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to pink.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xE1, 0x95, 0xC2);
            #endregion #TextConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "Quickly find books written by Ray Bradbury.";
        }
            finally
            {
                workbook.EndUpdate();
            }
        }


        static void AddSpecialConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #SpecialConditionalFormatting
            // Create the rule to identify unique values in cells A2 through A15. 
            SpecialConditionalFormatting cfRule = worksheet.ConditionalFormattings.AddSpecialConditionalFormatting(worksheet.Range["$A$2:$A$15"], ConditionalFormattingSpecialCondition.ContainUniqueValue);
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to yellow.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xFA, 0xF7, 0xAA);
            #endregion #SpecialConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "In a list of authors quickly identify unique values.";
         }
            finally
            {
                workbook.EndUpdate();
            }
        }


        static void AddTimePeriodConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfTasks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #TimePeriodConditionalFormatting
            // Create the rule to highlight today's dates in cells B2 through B6.
            TimePeriodConditionalFormatting cfRule =
            worksheet.ConditionalFormattings.AddTimePeriodConditionalFormatting(worksheet.Range["$B$2:$B$6"], ConditionalFormattingTimePeriod.Today);
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to pink.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xF2, 0xAE, 0xE3);
            #endregion #TimePeriodConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A8:B9"];
            ruleExplanation.Value = "Determine the today's task in the TO DO list.";
        }
            finally
            {
                workbook.EndUpdate();
            }
        }


        static void AddExpressionConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #ExpressionConditionalFormatting
            // Create the rule to identify values that are above the average in cells F2 through F15.
            ExpressionConditionalFormatting cfRule =
            worksheet.ConditionalFormattings.AddExpressionConditionalFormatting(worksheet.Range["$F$2:$F$15"], ConditionalFormattingExpressionCondition.GreaterThan, "=AVERAGE($F$2:$F$15)");
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to yellow.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xFA, 0xF7, 0xAA);
            // Set the font color to red.
            cfRule.Formatting.Font.Color = Color.Red;
            #endregion #ExpressionConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "Identify book prices that are greater than the average price.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }


        static void AddFormulaExpressionConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate(); 
            workbook.BeginUpdate();
            try
            {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            #region #FormulaExpressionConditionalFormatting
            // Create the rule to shade alternate rows without applying a new style.
            FormulaExpressionConditionalFormatting cfRule = worksheet.ConditionalFormattings.AddFormulaExpressionConditionalFormatting(worksheet.Range["$A$2:$G$15"], "=MOD(ROW(),2)=1");
            // Specify formatting options to be applied to cells if the condition is true.
            // Set the background color to light blue.
            cfRule.Formatting.Fill.BackgroundColor = Color.FromArgb(255, 0xBC, 0xDA, 0xF7);
            #endregion #FormulaExpressionConditionalFormatting
            // Add an explanation to the created rule.
            CellRange ruleExplanation = worksheet.Range["A17:G18"];
            ruleExplanation.Value = "Shade alternate rows in light blue without applying a new style.";
         }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddColorScale2ConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                workbook.Worksheets.ActiveWorksheet = worksheet;
                #region #ColorScale2ConditionalFormatting
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the minimum threshold to the lowest value in the range of cells.
                ConditionalFormattingValue minPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Set the maximum threshold to the highest value in the range of cells.
                ConditionalFormattingValue maxPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Create the two-color scale rule to differentiate low and high values in cells C2 through D15. Blue represents the lower values and yellow represents the higher values. 
                ColorScale2ConditionalFormatting cfRule = conditionalFormattings.AddColorScale2ConditionalFormatting(worksheet.Range["$C$2:$D$15"], minPoint, Color.FromArgb(255, 0x9D, 0xE9, 0xFA), maxPoint, Color.FromArgb(255, 0xFF, 0xF6, 0xA9));
                #endregion #ColorScale2ConditionalFormatting
                // Add an explanation to the created rule.
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Examine cost distribution using a gradation of two colors. Blue represents the lower values and yellow represents the higher values.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddColorScale3ConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                workbook.Worksheets.ActiveWorksheet = worksheet;
                #region #ColorScale3ConditionalFormatting
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the minimum threshold to the lowest value in the range of cells using the MIN() formula.
                ConditionalFormattingValue minPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Formula, "=MIN($C$2:$D$15)");
                // Set the midpoint threshold to the 50th percentile.
                ConditionalFormattingValue midPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percentile, "50");
                // Set the maximum threshold to the highest value in the range of cells using the MAX() formula.
                ConditionalFormattingValue maxPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Number, "=MAX($C$2:$D$15)");
                // Create the three-color scale rule to determine how values in cells C2 through D15 vary. Red represents the lower values, yellow represents the medium values and sky blue represents the higher values.
                ColorScale3ConditionalFormatting cfRule = conditionalFormattings.AddColorScale3ConditionalFormatting(worksheet.Range["$C$2:$D$15"], minPoint, Color.Red, midPoint, Color.Yellow, maxPoint, Color.SkyBlue);
                #endregion #ColorScale3ConditionalFormatting
                // Add an explanation to the created rule.
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Examine cost distribution using a gradation of three colors. Red represents the lower values, yellow represents the medium values and sky blue represents the higher values.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddDataBarConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                workbook.Worksheets.ActiveWorksheet = worksheet;
                #region #DataBarConditionalFormatting
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the value corresponding to the shortest bar to the lowest value.
                ConditionalFormattingValue lowBound1 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Set the value corresponding to the longest bar to the highest value.
                ConditionalFormattingValue highBound1 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Create the rule to compare values in cells E2 through E15 using data bars. 
                DataBarConditionalFormatting cfRule1 = conditionalFormattings.AddDataBarConditionalFormatting(worksheet.Range["$E$2:$E$15"], lowBound1, highBound1, DXColor.Green);
                // Set the positive bar border color to green.
                cfRule1.BorderColor = DXColor.Green;
                // Set the negative bar color to red.
                cfRule1.NegativeBarColor = DXColor.Red;
                // Set the negative bar border color to red.
                cfRule1.NegativeBarBorderColor = DXColor.Red;
                // Set the axis position to display the axis in the middle of the cell.
                cfRule1.AxisPosition = ConditionalFormattingDataBarAxisPosition.Middle;
                // Set the axis color to dark blue.
                cfRule1.AxisColor = Color.DarkBlue;

                // Set the value corresponding to the shortest bar to 0 percent.
                ConditionalFormattingValue lowBound2 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percent, "0");
                // Set the value corresponding to the longest bar to 100 percent.
                ConditionalFormattingValue highBound2 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percent, "100");
                // Create the rule to compare values in cells G2 through G15 using data bars.  
                DataBarConditionalFormatting cfRule2 = conditionalFormattings.AddDataBarConditionalFormatting(worksheet.Range["$G$2:$G$15"], lowBound2, highBound2, DXColor.SkyBlue);
                // Set the data bar border color to sky blue.
                cfRule2.BorderColor = DXColor.SkyBlue;
                // Specify the solid fill type.
                cfRule2.GradientFill = false;
                // Hide values of cells to which the rule is applied.
                cfRule2.ShowValue = false;
                #endregion #DataBarConditionalFormatting
                // Add an explanation to the created rule.
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Compare values in the \"Cost Trend\" and \"Markup\" columns using data bars.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }
        static void AddIconSetConditionalFormatting(IWorkbook workbook)
        {
            workbook.Calculate();
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                workbook.Worksheets.ActiveWorksheet = worksheet;
                #region #IconSetConditionalFormatting
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the first threshold to the lowest value in the range of cells using the MIN() formula.
                ConditionalFormattingIconSetValue minPoint = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Formula, "=MIN($E$2:$E$15)", ConditionalFormattingValueOperator.GreaterOrEqual);
                // Set the second threshold to 0.
                ConditionalFormattingIconSetValue midPoint = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Number, "0", ConditionalFormattingValueOperator.GreaterOrEqual);
                // Set the third threshold to 0.01.
                ConditionalFormattingIconSetValue maxPoint = conditionalFormattings.CreateIconSetValue(ConditionalFormattingValueType.Number, "0.01", ConditionalFormattingValueOperator.GreaterOrEqual);
                // Create the rule to apply a specific icon from the three arrow icon set to each cell in the range  E2:E15 based on its value.  
                IconSetConditionalFormatting cfRule = conditionalFormattings.AddIconSetConditionalFormatting(worksheet.Range["$E$2:$E$15"], IconSetType.Arrows3, new ConditionalFormattingIconSetValue[] { minPoint, midPoint, maxPoint });
                // Specify the custom icon to be displayed if the second condition is true. 
                // To do this, set the IconSetConditionalFormatting.IsCustom property to true, which is false by default.
                cfRule.IsCustom = true;
                // Initialize the ConditionalFormattingCustomIcon object.
                ConditionalFormattingCustomIcon cfCustomIcon = new ConditionalFormattingCustomIcon();
                // Specify the icon set where you wish to get the icon. 
                cfCustomIcon.IconSet = IconSetType.TrafficLights13;
                // Specify the index of the desired icon in the set.
                cfCustomIcon.IconIndex = 1;
                // Add the custom icon at the specified position in the initial icon set.
                cfRule.SetCustomIcon(1, cfCustomIcon);
                // Hide values of cells to which the rule is applied.
                cfRule.ShowValue = false;
                #endregion #IconSetConditionalFormatting
                // Add an explanation to the created rule.
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Identify upward and downward cost trends.";
            }
            finally
            {
                workbook.EndUpdate();
            }
        }
    }
}
