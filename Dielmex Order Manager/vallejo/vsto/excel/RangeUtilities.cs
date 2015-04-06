using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.vallejo.vsto.excel
{
    class RangeUtilities
    {
        #region < SINGLETON >
        private RangeUtilities _singletonInstance;

        public RangeUtilities GetInstance()
        {
            if (_singletonInstance == null)
            {
                _singletonInstance = new RangeUtilities();
            }

            return _singletonInstance;
        }
        #endregion


        /// <summary>
        /// param name="colName"
        /// param name="sheet"
        /// returns last used row
        /// 
        /// 
        /// Search last used row from a especified column
        /// </summary>
        public int getLastUsedIndex(int colName, string sheet)
        {
            int i = 0;
            i =+ i * 10;

        


            return 1;

        }

        public int searchInColumn()
        {
            return 0;
        }

        public bool existInRange(object value, object range, string sheetName)
        {

            return true;
        }

    }
}
