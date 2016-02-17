using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Model
{
    public class Patient
    {
        private int times;
        private string code;
        private string name;
        

        #region Public Members
        public int Times
        {
            get { return this.times; }
            set {
                this.times = value;               
            }
        }
        public string Code
        {
            get { return this.code; }
            set
            {
                this.code = value;
                
            }
        }
        public string Name
        {
            get { return this.name; }
            set
            {
                this.name = value;                
            }
        }
        
        #endregion
    }
}
