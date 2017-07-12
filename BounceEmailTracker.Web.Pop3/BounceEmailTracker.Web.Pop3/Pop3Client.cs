using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace BounceEmailTracker.Web.Pop3
{
    public class Pop3Client
    {
        private Pop3Credential m_credential;

        private const int m_pop3port = 110;
        private const int MAX_BUFFER_READ_SIZE = 256;

        private long m_inboxPosition = 0;
        private long m_directPosition = -1;

        private Socket m_socket = null;

        private Pop3Message m_pop3Message = null;

        public Pop3Credential UserDetails
        {
            set { m_credential = value; }
            get { return m_credential; }
        }

        public string From
        {
            get { return m_pop3Message.From; }
        }

        public string To
        {
            get { return m_pop3Message.To; }
        }

        public string Subject
        {
            get { return m_pop3Message.Subject; }
        }

        public string Body
        {
            get { return m_pop3Message.Body; }
        }

        public IEnumerator MultipartEnumerator
        {
            get { return m_pop3Message.MultipartEnumerator; }
        }

        public bool IsMultipart
        {
            get { return m_pop3Message.IsMultipart; }
        }

        public long InboxPosition
        {
            get { return m_pop3Message.InboxPosition; }
        }


        public Pop3Client(string user, string pass, string server)
        {
            m_credential = new Pop3Credential(user, pass, server);
        }

        private Socket GetClientSocket()
        {
            Socket s = null;

            try
            {
                IPHostEntry hostEntry = null;

                // Get host related information.
                hostEntry = Dns.GetHostEntry(m_credential.Server);

                // Loop through the AddressList to obtain the supported 
                // AddressFamily. This is to avoid an exception that 
                // occurs when the host IP Address is not compatible 
                // with the address family 
                // (typical in the IPv6 case).

                foreach (IPAddress address in hostEntry.AddressList)
                {
                    IPEndPoint ipe = new IPEndPoint(address, m_pop3port);

                    Socket tempSocket =
                        new Socket(ipe.AddressFamily,
                        SocketType.Stream, ProtocolType.Tcp);

                    tempSocket.Connect(ipe);

                    if (tempSocket.Connected)
                    {
                        // we have a connection.
                        // return this socket ...
                        s = tempSocket;
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                throw new Pop3ConnectException(e.ToString());
            }

            // throw exception if can't connect ...
            if (s == null)
            {
                throw new Pop3ConnectException("Error : connecting to "
                    + m_credential.Server);
            }

            return s;
        }

        public static string SendReceive(Socket socket, string szCommand, bool bWaitEndPop3Message)
        {
            // send user
            Send(socket, szCommand);

            string s = Receive(socket, bWaitEndPop3Message);
#if DEBUG_POP3
            System.Diagnostics.Debug.WriteLine("Receive.  Message: " + s);
#endif

            return s;
        }

        private static void Send(Socket socket, string szCommand)
        {
            byte[] byteData = Encoding.ASCII.GetBytes(szCommand + "\r\n");
            socket.Send(byteData);

#if DEBUG_POP3
            System.Diagnostics.Debug.WriteLine("Send.  Message: " + szCommand);
#endif
        }

        private static string Receive(Socket socket, bool bWaitEndPop3Message)
        {
            System.Threading.Thread.Sleep(500);

            const int nMaxSize = 1000000;
            // wait timeout on receiving (in milliseconds)
            const int cnTimeout = 30000;
            byte[] bReceiveBytes = new byte[nMaxSize];

            // receive the response
            int offset = 0;
            int byteCount = 0;
            while (true)
            {
                System.Threading.Thread.Sleep(50);
                IAsyncResult asResult = socket.BeginReceive(bReceiveBytes, offset, nMaxSize - offset, SocketFlags.None, null, null);
                if (!asResult.AsyncWaitHandle.WaitOne(cnTimeout, false))
                    //throw new Exception("WaitHandle returned false in GetSocketData");
                    return "";
                byteCount += socket.EndReceive(asResult);

                // no bytes received
                if (0 == byteCount)
                    return "";

                offset = byteCount;
                if (offset >= nMaxSize)
                    break;

                string s = (new System.Text.ASCIIEncoding()).GetString(bReceiveBytes, 0, byteCount);
                if (s.EndsWith("\r\n.\r\n"))
                    break;

                if (!bWaitEndPop3Message)
                    break;
            }

            // get the message header
            string szResponse = (new System.Text.ASCIIEncoding()).GetString(bReceiveBytes, 0, byteCount);
            return szResponse;
        }

        private void LoginToInbox()
        {
            string returned = SendReceive(m_socket, "USER " + m_credential.User, false);

            if (!returned.Substring(0, 3).Equals("+OK"))
            {
                throw new Pop3LoginException("login not excepted");
            }

            returned = SendReceive(m_socket, "PASS " + m_credential.Pass, false);

            if (!returned.Substring(0, 3).Equals("+OK"))
            {
                throw new
                    Pop3LoginException("login/password not accepted");
            }
        }

        public long MessageCount
        {
            get
            {
                long count = 0;

                if (m_socket == null)
                {
                    throw new Pop3MessageException("Pop3 server not connected");
                }

                string returned = SendReceive(m_socket, "STAT", false);

                //Send("stat");

                //string returned = GetPop3String();

                // if values returned ...
                if (Regex.Match(returned,
                    @"^.*\+OK[ |	]+([0-9]+)[ |	]+.*$").Success)
                {
                    // get number of emails ...
                    count = long.Parse(Regex
                    .Replace(returned.Replace("\r\n", "")
                    , @"^.*\+OK[ |	]+([0-9]+)[ |	]+.*$", "$1"));
                }

                return (count);
            }
        }


        public void CloseConnection()
        {
            Send(m_socket, "quit");

            m_socket = null;
            m_pop3Message = null;
        }

        public bool DeleteEmail()
        {
            bool ret = false;

            string returned = SendReceive(m_socket, "dele " + m_inboxPosition, false);

            if (Regex.Match(returned,
                @"^.*\+OK.*$").Success)
            {
                ret = true;
            }

            return ret;
        }

        public bool NextEmail(long directPosition)
        {
            bool ret;

            if (directPosition >= 0)
            {
                m_directPosition = directPosition;
                ret = NextEmail();
            }
            else
            {
                throw new Pop3MessageException("Position less than zero");
            }

            return ret;
        }

        public bool NextEmail()
        {
            string returned;

            long pos;

            if (m_directPosition == -1)
            {
                //m_inboxPosition = 375;
                if (m_inboxPosition == 0)
                {
                    pos = 1;
                }
                else
                {
                    pos = m_inboxPosition + 1;
                }
            }
            else
            {
                pos = m_directPosition + 1;
                m_directPosition = -1;
            }

            // send list command...
            returned = SendReceive(m_socket, "LIST " + pos.ToString(), false);

            // if email does not exist at this position
            // then return false ...

            if (4 > returned.Length || returned.Substring(0, 4).Equals("-ERR"))
                return false;

            m_inboxPosition = pos;

            // strip out CRLF ...
            string[] noCr = returned.Split(new char[] { '\r' });
            if (null == noCr || 0 == noCr.Length)
                return false;

            // get size ...
            string[] elements = noCr[0].Split(new char[] { ' ' });
            if (null == elements || 3 > elements.Length)
                return false;

            long size = long.Parse(elements[2]);

            // ... else read email data
            m_pop3Message = new Pop3Message(m_inboxPosition, size, m_socket);

            return true;
        }

        public void OpenInbox()
        {
            // get a socket ...
            m_socket = GetClientSocket();

            // get initial header from POP3 server ...
            string header = Receive(m_socket, false);

            if (!header.Substring(0, 3).Equals("+OK"))
            {
                throw new Exception("Invalid initial POP3 response");
            }

            // send login details ...
            LoginToInbox();
        }
    }
}
