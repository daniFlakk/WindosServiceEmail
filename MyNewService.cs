using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace MyNewService
{
    public partial class MyNewService : ServiceBase
    {
        private int eventId = 1;

        private string _emailFrom;

        public object ConfigurationManager { get; private set; }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);
        public MyNewService()
        {

            InitializeComponent();

            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";
        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Set up a timer that triggers every minute.
            Timer timer = new Timer();
            timer.Interval = 60000; // 60 seconds
            
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
            eventLog1.WriteEntry("In OnStart.");


            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            using (var reader = new StreamReader(@"direccion donde esta el csv"))
            
            {

                while (!reader.EndOfStream)
                {

                    var line = reader.ReadLine();

                    var values = line.Split(',');


                    // Obtener la fecha especificada en el archivo
                    if (DateTime.TryParse(values[2], out DateTime sendDate))
                    {
                        eventLog1.WriteEntry("fecha?" + line);
                        // Verificar si la fecha especificada es igual a la del siguiente  dia
                        if (DateTime.Now.Date.AddDays(1) == sendDate.Date && values[3].Equals("pendiente"))
                        {
                            try
                            {
                                MailMessage correo = new MailMessage();
                                correo.From = new MailAddress("correo del que se enviaran", "Nombre que se le quiere dar a las notificaciones", System.Text.Encoding.UTF8);//Correo de salida
                                correo.To.Add(values[1]); //Correo destino? 
                                correo.Subject = "Recordatorio cita Dr. Vargas" + values[2]; //Asunto
                                string tomorrowDate = (DateTime.Today.AddDays(1)).ToString("dd/MM/yyyy");
                                correo.Body = "Estimado/a " + values[0] + " Le recordamos que el dia de mañana, " + tomorrowDate + " con el doctor Andres Camilo Vargas"; //Mensaje del correo
                                correo.IsBodyHtml = true;
                                correo.Priority = MailPriority.Normal;
                                SmtpClient smtp = new SmtpClient();
                                smtp.UseDefaultCredentials = false;
                                smtp.Host = "smtp.office365.com"; //Host del servidor de correo en este caso es de outlook
                                smtp.Port = 587; //Puerto de salida
                                smtp.Credentials = new System.Net.NetworkCredential("correo de salidad", "contraseña");//Cuenta de correo
                                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                                smtp.EnableSsl = true;//True si el servidor de correo permite ssl
                                smtp.Send(correo);
                                eventLog1.WriteEntry("se envio " + values[1]);
                                //ActualizarEstado(@"C:\Users\Santiago Ortega\Desktop\educacion\Universidad\Semestre 9\Sistemas operativos\users.csv", line, "enviado");
                            }
                            catch (Exception ex)
                            {
                                eventLog1.WriteEntry("Ha ocurrido un error al enviar el correo: " + ex.Message + values[1]);
                            }

                        }
                    }

                }
            }

    } 
        



        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");
        }


        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };


        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }


    }
}
