using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookAndFeel.Infraestructure
{
    using ViewModels;
    public class InstanceLocatior
    {

        #region MyProperty
        public MainViewModel main { get; set; }
        #endregion
        #region Constructors
        public InstanceLocatior()
        {
            this.main = new MainViewModel();
        }
        #endregion

    }
}
