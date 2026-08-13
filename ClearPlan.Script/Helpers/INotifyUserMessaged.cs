using System;

namespace ClearPlan
{
    public interface INotifyUserMessaged
    {
        event EventHandler<UserMessagedEventArgs> UserMessaged;
    }
}
