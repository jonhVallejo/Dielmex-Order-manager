using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    class Service
    {
        internal string ServiceId
        {
            get;
            set;
        }

        internal string Description
        {
            get;
            set;
        }

        internal string UnitOfMeasurement
        {
            get;
            set;
        }

        

        internal double Cost
        {
            get;
            set;
        }


        public override string ToString()
        {
            return ServiceId;
        }


    }
}
