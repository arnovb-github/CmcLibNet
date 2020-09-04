using System.Text;

namespace Vovin.CmcLibNet.Database
{
    // contains the methods to create a [ViewFilter(bla)] DDE request string from the filter objects
    internal static class FilterFormatters
    {
        internal static string FormatCTIFilter(ICursorFilter filter)
        {
            ICursorFilterTypeCTI f = (CursorFilterTypeCTI)filter;
            StringBuilder sb = new StringBuilder("[ViewFilter(");
            sb.Append(f.ClauseNumber.ToString());
            sb.Append(',');
            sb.Append(f.FiltertypeIdentifier);
            sb.Append(',');
            sb.Append(f.Except ? "NOT," : ",");
            sb.Append(Utils.Dq(f.Connection));
            sb.Append(',');
            sb.Append(Utils.Dq(f.Category));
            sb.Append(',');
            sb.Append(Utils.Dq(Utils.GetClarifiedItemName(f.Item, f.ClarifySeparator, f.ClarifyValue)));
            sb.Append(")]");
            return sb.ToString();
        }

        internal static string FormatFFilter(ICursorFilter filter)
        {
            ICursorFilterTypeF f = (ICursorFilterTypeF)filter;
            StringBuilder sb = new StringBuilder("[ViewFilter(");
            sb.Append(f.ClauseNumber.ToString());
            sb.Append(',');
            sb.Append(f.FiltertypeIdentifier);
            sb.Append(',');
            sb.Append((f.Except) ? "NOT," : ",");
            if (f.SharedOptionSet)
            {
                sb.Append("," + Utils.Dq((f.Shared) ? "Shared" : "Local"));
                sb.Append(",,"); // two!
            }
            else
            {
                sb.Append(Utils.Dq(f.FieldName));
                sb.Append(',');
                sb.Append(Utils.Dq(f.QualifierString));
                sb.Append(',');
                if (f.Qualifier == FilterQualifier.Between) // TODO code smell, this value may not be set by COM clients
                {
                    sb.Append(Utils.Dq(f.FilterBetweenStartValue));
                    sb.Append(',');
                    sb.Append(Utils.Dq(f.FilterBetweenEndValue));
                }
                else
                {
                    sb.Append(Utils.Dq(f.FieldValue));
                    sb.Append(',');
                    sb.Append((f.MatchCase) ? "1" : "0");
                }
            }
            sb.Append(")]");
            return sb.ToString();
        }

        internal static string FormatCTCTIFilter(ICursorFilter filter)
        {
            ICursorFilterTypeCTCTI f = (CursorFilterTypeCTCTI)filter;
            StringBuilder sb = new StringBuilder("[ViewFilter(");
            sb.Append(f.ClauseNumber.ToString());
            sb.Append(',');
            sb.Append(f.FiltertypeIdentifier);
            sb.Append(',');
            sb.Append((f.Except) ? "NOT," : ",");
            sb.Append(Utils.Dq(f.Connection));
            sb.Append(',');
            sb.Append(Utils.Dq(f.Category));
            sb.Append(',');
            sb.Append(Utils.Dq(f.Connection2));
            sb.Append(',');
            sb.Append(Utils.Dq(f.Category2));
            sb.Append(',');
            sb.Append(Utils.Dq(Utils.GetClarifiedItemName(f.Item, f.ClarifySeparator, f.ClarifyValue)));
            sb.Append(")]");
            return sb.ToString();
        }

        internal static string FormatCTCFFilter(ICursorFilter filter)
        {
            ICursorFilterTypeCTCF f = (CursorFilterTypeCTCF)filter;
            StringBuilder sb = new StringBuilder("[ViewFilter(");
            sb.Append(f.ClauseNumber.ToString());
            sb.Append(',');
            sb.Append(f.FiltertypeIdentifier);
            sb.Append(',');
            sb.Append(f.Except ? "NOT," : ",");
            sb.Append(Utils.Dq(f.Connection));
            sb.Append(',');
            sb.Append(Utils.Dq(f.Category));
            sb.Append(',');
            if (f.SharedOptionSet)
            {
                sb.Append("," + Utils.Dq((f.Shared) ? "Shared" : "Local"));
                sb.Append(",,"); // two!
            }
            else
            {
                sb.Append(Utils.Dq(f.FieldName));
                sb.Append(',');
                sb.Append(Utils.Dq(f.QualifierString));
                sb.Append(',');
                if (f.Qualifier == FilterQualifier.Between)  // TODO code smell, this value may not be set by COM clients
                {
                    sb.Append(Utils.Dq(f.FilterBetweenStartValue));
                    sb.Append(',');
                    sb.Append(Utils.Dq(f.FilterBetweenEndValue));
                }
                else
                {
                    sb.Append(Utils.Dq(f.FieldValue));
                    sb.Append(',');
                    sb.Append((f.MatchCase) ? "1" : "0");
                }
            }
            sb.Append(")]");
            return sb.ToString();
        }
    }
}
