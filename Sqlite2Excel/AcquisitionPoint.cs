using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sqlite2Excel
{
    public class AcquisitionPoint
    {
        private string id;
        private string stakeId;
        private string directionType;

        public AcquisitionPoint(string id, string stakeId, string type)
        {
            this.id = id;
            this.stakeId = stakeId;
            this.directionType = type;
        }

        public string GetId()
        {
            return this.id;
        }

        public string GetStakeId()
        {
            return this.stakeId;
        }

        public string GetDirectionType()
        {
            return this.directionType;
        }
    }
}