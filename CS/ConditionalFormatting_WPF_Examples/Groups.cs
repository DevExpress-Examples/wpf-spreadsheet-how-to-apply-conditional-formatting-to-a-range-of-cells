using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConditionalFormatting_WPF_Examples
{
    public partial class Groups : List<Group>
    {
        public static Groups InitData()
        {
            Groups examples = new Groups();

            #region GroupNodes
            examples.Add(new Group("Conditional Formatting Examples"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Conditional Formatting" group of examples.
            examples[0].Items.Add(new SpreadsheetExample("Format values that are above or below average", ConditionalFormatting.AddAverageConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells that are not between two specified values", ConditionalFormatting.AddRangeConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format top ranked values", ConditionalFormatting.AddRankConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells that contain the given text", ConditionalFormatting.AddTextConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format only unique values", ConditionalFormatting.AddSpecialConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format today's date", ConditionalFormatting.AddTimePeriodConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format values that are greater than a specified value", ConditionalFormatting.AddExpressionConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Use a formula to determine which cells to format", ConditionalFormatting.AddFormulaExpressionConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells by using a two-color scale", ConditionalFormatting.AddColorScale2ConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells by using a three-color scale", ConditionalFormatting.AddColorScale3ConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells by using data bars", ConditionalFormatting.AddDataBarConditionalFormattingAction));
            examples[0].Items.Add(new SpreadsheetExample("Format cells by using a custom icon set", ConditionalFormatting.AddIconSetConditionalFormattingAction));

            return examples;
            #endregion
        }
    }
}
