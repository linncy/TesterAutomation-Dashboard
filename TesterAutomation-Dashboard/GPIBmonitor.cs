using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ivi.Visa;
using Ivi.Visa.FormattedIO;

namespace TesterAutomation_Dashboard
{
    public class GPIBmonitor:Form1
    {
        private double CpMonitor(IMessageBasedSession session, MessageBasedFormattedIO formattedIO)
        {
            return 1.0;
        }
        private double GpMonitor(IMessageBasedSession session, MessageBasedFormattedIO formattedIO)
        {
            return 1.0;
        }
        private double RpMonitor(IMessageBasedSession session, MessageBasedFormattedIO formattedIO)
        {
            return 1.0;
        }
        private bool CGRmonitor(IMessageBasedSession session, MessageBasedFormattedIO formattedIO)
        {
            MeaLabelModify(1, 2.123);
            return true;
        }
    }
}
