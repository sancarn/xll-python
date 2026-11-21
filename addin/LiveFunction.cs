using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ExcelDna.Integration;

namespace PythonExcel
{
    internal sealed class LiveFunction : IExcelObservable
    {
        sealed class Subscription : IDisposable
        {
            internal readonly string Id = Guid.NewGuid().ToString("N");
            internal readonly string Name;
            internal readonly IExcelObserver Observer;

            internal Subscription(string name, IExcelObserver observer)
            {
                Name = name;
                Observer = observer;
            }

            public void Dispose()
            {
                lock (subscriptions)
                {
                    if (subscriptions.Remove(Id))
                    {
                        try { Runtime.Request(new { op="streamClose", id=Id }); }
                        catch (Exception ex) { Runtime.Log(Name + "\n" + ex); }
                    }
                }
            }
        }

        static readonly Dictionary<string, Subscription> subscriptions = new Dictionary<string, Subscription>();
        static Thread reader;
        readonly string name;
        readonly object[] arguments;

        internal LiveFunction(string name, object[] arguments)
        {
            this.name = name;
            this.arguments = arguments;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            var subscription = new Subscription(name, observer);
            try
            {
                lock (subscriptions)
                {
                    Runtime.Request(new { op="streamOpen", id=subscription.Id, name, args=arguments });
                    subscriptions.Add(subscription.Id, subscription);
                    if (reader == null)
                    {
                        var thread = new Thread(ReadUpdates) { IsBackground=true, Name="Python worksheet updates" };
                        thread.Start();
                        reader = thread;
                    }
                }
            }
            catch (Exception ex)
            {
                subscription.Dispose();
                Runtime.Log(name + "\n" + ex);
                observer.OnNext(Tuple.Create((object)ExcelError.ExcelErrorValue));
            }
            return subscription;
        }

        static void ReadUpdates()
        {
            while (true)
            {
                lock (subscriptions)
                {
                    if (subscriptions.Count == 0)
                    {
                        reader = null;
                        return;
                    }
                }
                IDictionary update;
                try { update = (IDictionary)Runtime.Request(new { op="streamRead" }); }
                catch (Exception ex)
                {
                    Runtime.Log("Reading Python worksheet updates\n" + ex);
                    Subscription[] failed;
                    lock (subscriptions)
                    {
                        failed = subscriptions.Values.ToArray();
                        DisconnectAll();
                        reader = null;
                    }
                    foreach (var topic in failed) Publish(topic, ExcelError.ExcelErrorValue);
                    return;
                }
                if (update == null) continue;
                Subscription sub;
                lock (subscriptions)
                    if (!subscriptions.TryGetValue((string)update["id"], out sub)) continue;
                object value;
                try
                {
                    if (update.Contains("failure")) throw new InvalidOperationException((string)update["failure"]);
                    value = Functions.ToExcel(Runtime.Json.DeserializeObject((string)update["json"]));
                }
                catch (Exception ex)
                {
                    Runtime.Log(sub.Name + "\n" + ex.Message);
                    value = ExcelError.ExcelErrorValue;
                }
                // Keep the final value subscribed until Excel drops the topic, including across recalculations.
                Publish(sub, value);
            }
        }

        static void Publish(Subscription sub, object value)
        {
            try { sub.Observer.OnNext(Tuple.Create(value)); }
            catch (Exception ex) { Runtime.Log(sub.Name + "\n" + ex); }
        }

        internal static void DisconnectAll()
        {
            lock (subscriptions)
                foreach (var sub in subscriptions.Values.ToArray()) sub.Dispose();
        }
    }
}
