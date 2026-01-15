using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DonateMonitor.ServiceListener
{
    internal interface IServiceListener
    {
        Task StartAsync(Monitor monitor, CancellationToken token);
        Task StopAsync();
    }
}
