namespace Vovin.CmcLibNet.Services
{
    // POCO class to hold info on Commence fields
    internal class Field
    {
        #region Constructors
        internal Field() { }

        internal Field(string fieldName, string fieldLabel, string fieldValue)
        {
            Name = fieldName;
            Label = fieldLabel;
            Value = fieldValue;
        }
        #endregion

        #region Properties
        public string Value { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;
        #endregion
    }
}
