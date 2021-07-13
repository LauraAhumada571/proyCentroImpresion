using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace proyCentroImpresion
{
    public abstract class ViewModelBase : INotifyPropertyChanging, INotifyPropertyChanged

    {
        public event PropertyChangingEventHandler PropertyChanging;
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanging([CallerMemberName] string propertyName = "")
        {
            PropertyChanging(this, new PropertyChangingEventArgs(propertyName));
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
