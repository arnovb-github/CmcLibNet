using System;
using System.Linq;
using System.Reflection;
using Vovin.CmcLibNet.Database;

namespace Vovin.CmcLibNet.Export.Complex
{

    internal class CursorFactory : IDisposable
    {
        private readonly ICommenceDatabase db;

        internal CursorFactory()
        {
            db = new CommenceDatabase();
        }

        internal ICommenceCursor Create(CursorDescriptor cursorParameters)
        {
            ICommenceCursor retval = null;
            if (db == null)
            {
                throw new InvalidOperationException($"Underlying database was closed.");
            }
            retval = db.GetCursor(cursorParameters.CategoryOrView, cursorParameters.CursorType, CmcOptionFlags.UseThids | CmcOptionFlags.IgnoreSyncCondition);
            retval.SetColumns(cursorParameters.Fields.ToArray());
            retval.Columns.Apply();
            for (int i = 0; i < cursorParameters.Filters.Count(); i++)
            {
                ICursorFilterTypeCTCF f = retval.Filters.Add(i + 1, FilterType.ConnectionToCategoryField);
                // we already have the filter object defined, copy them
                foreach (PropertyInfo property in typeof(ICursorFilterTypeCTCF).GetProperties().Where(p => p.CanWrite))
                {
                    property.SetValue(f, property.GetValue(cursorParameters.Filters[i], null), null);
                }
            }
            retval.Filters.Apply();
            retval.MaxFieldSize = cursorParameters.MaxFieldSize;
            return retval;
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    db.Close();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~CursorFactory() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
