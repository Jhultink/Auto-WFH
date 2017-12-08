using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AutoWFH
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }

        static async Task MainAsync(string[] args)
        {
            string logPath = "log.txt";
            using (StreamWriter log = new StreamWriter(logPath, true))
            {
                
                log.WriteLine("****** " + DateTime.Now.ToLocalTime() + " ******");

                var settings = new XmlDocument();
                settings.Load("Settings.config");
                XmlNode root = settings.SelectSingleNode("Settings");

                string steamKey = root.SelectSingleNode("SteamKey").InnerText;
                string xboxKey = root.SelectSingleNode("XboxKey").InnerText;

                HttpClient steamClient = new HttpClient();
                HttpClient xboxClient = new HttpClient();
                xboxClient.DefaultRequestHeaders.Add("X-AUTH", xboxKey);

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = "smtp.bdo.com";

                List<string> emailList = new List<string>();
                foreach (XmlNode emailNode in root.SelectSingleNode("EmailList"))
                {
                    if (emailNode.NodeType != XmlNodeType.Comment)
                    {
                        emailList.Add(emailNode.InnerText);
                    }
                }

                foreach (XmlNode userNode in root.SelectSingleNode("Users").ChildNodes)
                {
                    if (userNode.NodeType != XmlNodeType.Comment)
                    {
                        string name = userNode.SelectSingleNode("Name").InnerText;
                        string steamId = userNode.SelectSingleNode("SteamId").InnerText;
                        string email = userNode.SelectSingleNode("Email").InnerText;
                        string cutoffStr = userNode.SelectSingleNode("WFHCutoff").InnerText;
                        string signature = userNode?.SelectSingleNode("Signature")?.InnerXml ?? "";
                        string excuse = "";
                        DateTime cutoff = DateTime.ParseExact(cutoffStr, "h:mm tt", CultureInfo.InvariantCulture);

                        // Get random excuse
                        XmlNode excusesNode = userNode.SelectSingleNode("Excuses");
                        if ((excusesNode?.ChildNodes?.Count ?? -1) > 0)
                        {
                            int excuseIndex = new Random().Next(excusesNode.ChildNodes.Count);
                            excuse = excusesNode.ChildNodes[excuseIndex].InnerText;
                        }
                        else
                        {
                            excuse = root.SelectSingleNode("DefaultExcuse").InnerText;
                        }

                        DateTime lastSteamLogoff = DateTime.MinValue;

                        try
                        {
                            // Get Steam last logoff
                            string resp = await steamClient.GetStringAsync("http://api.steampowered.com/ISteamUser/GetPlayerSummaries/v0002/?key=" + steamKey + "&steamids=" + steamId);
                            SteamUser user = JsonConvert.DeserializeObject<SteamUser>(resp);
                            int unixLoggoff = user.response.players.First().lastlogoff;
                            lastSteamLogoff = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(unixLoggoff).ToLocalTime();

                            log.WriteLine(name + " logged off Steam at " + lastSteamLogoff);
                        }
                        catch(Exception e)
                        {
                            log.WriteLine("Error pulling Steam data");
                            log.WriteLine(e.ToString());
                        }

                        // Get Xbox last logoff
                        DateTime xboxLogoff = DateTime.MinValue;
                        try
                        {
                            string xboxId = userNode?.SelectSingleNode("XboxId")?.InnerXml ?? "";
                            if (!String.IsNullOrEmpty(xboxId))
                            {
                                HttpResponseMessage xboxResp = await xboxClient.GetAsync("https://xboxapi.com/v2/" + xboxId + "/presence");
                                string xboxRespStr = await xboxResp.Content.ReadAsStringAsync();
                                JObject xboxRespObj = JObject.Parse(xboxRespStr);

                                if (xboxRespObj["state"].ToString() == "Online")
                                {
                                    xboxLogoff = DateTime.Now;
                                }
                                else if (xboxRespObj["state"].ToString() == "Offline")
                                {
                                    var lastSeen = xboxRespObj["lastSeen"];
                                    if (lastSeen != null && lastSeen["deviceType"].ToString() == "XboxOne")
                                    {
                                        string xboxTimeString = lastSeen["timestamp"].ToString();
                                        xboxLogoff = DateTime.Parse(xboxTimeString);
                                    }
                                }

                            }
                        }
                        catch(Exception e)
                        {
                            log.WriteLine("Error pulling Xbox data");
                            log.WriteLine(e.ToString());
                        }

                        string platform = "Steam";

                        DateTime lastActivity = lastSteamLogoff;
                        if (xboxLogoff > lastActivity)
                        {
                            lastActivity = xboxLogoff;
                            log.WriteLine(name + " logged off Xbox at " + xboxLogoff);
                            platform = "Xbox";
                        }


                        if (lastActivity > cutoff)
                        {
                            try
                            {

                                MailMessage msg = new MailMessage()
                                {
                                    From = new MailAddress(email),
                                    Subject = "WFH",
                                    Body = excuse + " <p style=\"color:blue; font-size:250% \"> Courtesy of Steam Catcher for " + platform + " </p> " + signature,
                                    IsBodyHtml = true
                                };

                                emailList.Where(x => !x.Equals(email, StringComparison.CurrentCultureIgnoreCase)).ToList().ForEach(m => msg.To.Add(m));
#if DEBUG
                                msg.CC.Add(new MailAddress("Jhultink@bdo.com"));
#else
                                msg.CC.Add(new MailAddress(email));
#endif
                                log.WriteLine("Sent email for " + name + " on " + platform);
                            }
                            catch (Exception e)
                            {
                                log.WriteLine("Error sending email for " + name);
                                log.WriteLine(e.ToString());
                            }
                            await smtpClient.SendMailAsync(msg);
                        }
                    }
                }
            }
        }
    }

    public class Player
    {
        public string steamid { get; set; }
        public int communityvisibilitystate { get; set; }
        public int profilestate { get; set; }
        public string personaname { get; set; }
        public int lastlogoff { get; set; }
        public string profileurl { get; set; }
        public string avatar { get; set; }
        public string avatarmedium { get; set; }
        public string avatarfull { get; set; }
        public int personastate { get; set; }
        public string realname { get; set; }
        public string primaryclanid { get; set; }
        public int timecreated { get; set; }
        public int personastateflags { get; set; }
    }

    public class Response
    {
        public List<Player> players { get; set; }
    }

    public class SteamUser
    {
        public Response response { get; set; }
    }
}
