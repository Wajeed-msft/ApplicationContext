using Newtonsoft.Json;
using System.Collections.Generic;

namespace TeamsAdmin.Models
{
    public class GroupInfo
    {
        public string description { get; set; }
        public string displayName { get; set; }
        public string[] groupTypes { get; set; }
        public bool mailEnabled { get; set; }
        public string mailNickname { get; set; }
        public bool securityEnabled { get; set; }
        [JsonProperty("owners@odata.bind")]
        public string[] ownersodatabind { get; set; }
    }

    #region Team Creation POCOs
    public class TeamSettings
    {
        public Membersettings memberSettings { get; set; }
        public Messagingsettings messagingSettings { get; set; }
        public Funsettings funSettings { get; set; }
    }

    public class Membersettings
    {
        public bool allowCreateUpdateChannels { get; set; }
    }

    public class Messagingsettings
    {
        public bool allowUserEditMessages { get; set; }
        public bool allowUserDeleteMessages { get; set; }
    }

    public class Funsettings
    {
        public bool allowGiphy { get; set; }
        public string giphyContentRating { get; set; }
    }
    #endregion

    public class ChannelInfoBody
    {
        public string displayName { get; set; }
        public string description { get; set; }
    }

    public class NewTeamDetails
    {
        public string OwnerEmailId { get; set; }
        public string TeamName { get; set; }
        public List<string> ChannelNames { get; set; } = new List<string>();
        public List<string> MemberEmails { get; set; } = new List<string>();
    }

    public class ResponseData
    {
        public string id { get; set; }
    }

}