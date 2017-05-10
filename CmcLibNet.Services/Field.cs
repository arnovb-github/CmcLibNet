namespace Vovin.CmcLibNet.Services
{
    // POCO class to hold info on Commence fields
    internal class Field
    {
        private string _name = string.Empty;
        private string _label = string.Empty;
        private string _value = string.Empty;

        #region Conctructors
        internal Field() { }

        internal Field(string fieldName, string fieldLabel, string fieldValue)
        {
            _name = fieldName;
            _label = fieldLabel;
            _value = fieldValue;
        }
        #endregion

        #region Properties
        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                _value = value;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }

        public string Label
        {
            get
            {
                return _label;
            }
            set
            {
                _label = value;
            }
        }
        #endregion
    }
}
