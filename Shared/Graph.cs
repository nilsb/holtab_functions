using Shared.Models;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Logging;
using StackExchange.Redis;
using Newtonsoft.Json;
using System.Reflection.Metadata;
using System.Linq;
using System.Dynamic;

namespace Shared
{
    public class Graph
    {
        private readonly GraphServiceClient? graphClient;
        private readonly ILogger? log;
        private readonly Services? services;
        private readonly Settings? settings;
        private readonly int maxUploadChunkSize = 320 * 1024;
        private readonly string? SqlConnectionString;
        private readonly string? redisConnectionString;
        private readonly ConnectionMultiplexer? redis;
        private readonly IDatabase? redisDB;

        public Graph(Settings _settings)
        {
            if (_settings != null)
            {
                settings = _settings;
                graphClient = _settings.GraphClient;
                log = _settings.log;
                SqlConnectionString = _settings.SqlConnectionString;
                redisConnectionString = _settings.redisConnectionString;

                if (!string.IsNullOrEmpty(SqlConnectionString))
                {
                    services = new Services(SqlConnectionString, log);
                }

                if (!string.IsNullOrEmpty(redisConnectionString))
                {
                    redis = ConnectionMultiplexer.Connect(redisConnectionString);

                    if(redis != null)
                    {
                        redisDB = redis.GetDatabase();
                    }
                }
            }
        }

        #region Users
        public async Task<User?> GetUser(string userEmail)
        {
            User? returnValue = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                log?.LogInformation($"Trying to find member {userEmail}");

                returnValue = await graphClient.Users[userEmail].GetAsync();
            }
            catch (Exception)
            {
                log?.LogInformation($"Unable to find member {userEmail}");
            }

            return returnValue;
        }

        #endregion

        #region Teams
        public async Task<Chat?> CreateOneOnOneChat(List<string> members)
        {
            Chat? returnValue = null;

            try
            {
                List<ConversationMember> chatMembers = new List<ConversationMember>();

                foreach (string member in members)
                {
                    User? memberUser = await GetUser(member);

                    if(memberUser != null)
                    {
                        chatMembers.Add(new AadUserConversationMember
                        {
                            Roles = new List<string>
                        {
                            "owner"
                        },
                            AdditionalData = new Dictionary<string, object>
                        {
                            { "user@odata.bind", "https://graph.microsoft.com/v1.0/users('" + memberUser.Id + "')" }
                        }
                        });
                    }
                }

                var chat = new Chat()
                {
                    ChatType = ChatType.OneOnOne,
                    Members = chatMembers
                };

                if(graphClient != null)
                {
                    returnValue = await graphClient.Chats.PostAsync(chat);
                }
            }
            catch (Exception)
            {
            }

            return returnValue;
        }

        public async Task<ChatMessage?> SendOneOnOneMessage(Chat existingChat, string content)
        {
            ChatMessage? returnValue = null;

            try
            {
                var chatMessage = new ChatMessage()
                {
                    Body = new ItemBody()
                    {
                        Content = content
                    }
                };

                if(graphClient != null)
                {
                    returnValue = await graphClient.Chats[existingChat.Id].Messages.PostAsync(chatMessage);
                }
            }
            catch (Exception)
            {
            }

            return returnValue;
        }

        public async Task<bool> AddTeamMember(string userEmail, string TeamId, string role)
        {
            User? memberToAdd = default(User);
            bool returnValue = false;

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null && redis != null && redisDB != null)
            {
                RedisValue cachedValue = redisDB.StringGet(userEmail);

                if (cachedValue.IsNullOrEmpty)
                {
                    try
                    {
                        log?.LogInformation($"Trying to find member {userEmail}");
                        memberToAdd = await graphClient.Users[userEmail].GetAsync();
                    }
                    catch (Exception)
                    {
                        log?.LogInformation($"Unable to find member {userEmail}");
                    }
                }

                if (cachedValue.IsNullOrEmpty)
                {
                    cachedValue = redisDB.StringSetAndGet(userEmail, memberToAdd?.Id);
                }

                if (cachedValue.HasValue)
                {
                    log?.LogInformation($"Adding member {userEmail} to team");

                    var conversationMember = new AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {
                            role
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('" + cachedValue.ToString() + "')"}
                        }
                    };

                    try
                    {
                        await graphClient.Teams[TeamId].Members.PostAsync(conversationMember);
                        returnValue = true;
                    }
                    catch (Exception ex)
                    {
                        log?.LogError(ex.Message);
                    }
                }
            }

            return returnValue;
        }

        //public async Task<Team?> GetTeamFromGroup(Group group)
        //{
        //    Team? foundTeam = null;
        //    RedisValue cachedValue = redisDB.StringGet($"Team for: {group.Id}");

        //    if (!cachedValue.IsNullOrEmpty)
        //        return JsonConvert.DeserializeObject<Team>(cachedValue);

        //    if (graphClient == null)
        //    {
        //        return foundTeam;
        //    }

        //    try
        //    {
        //        foundTeam = await graphClient.Groups[group.Id].Team.GetAsync();
        //    }
        //    catch (Exception)
        //    {
        //    }

        //    redisDB.StringSet($"Team for: {group.Id}", JsonConvert.SerializeObject(foundTeam));

        //    return foundTeam;
        //}

        public async Task<string?> GetTeamFromGroup(string groupId, bool debug)
        {
            string? foundTeam = null;

            if(redisDB == null || graphClient == null)
            {
                return foundTeam;
            }

            RedisValue cachedValue = redisDB.StringGet($"TeamId for: {groupId}");

            if (cachedValue.HasValue && !cachedValue.IsNullOrEmpty)
            {
                if(debug)
                    log?.LogInformation($"GetTeamFromGroup: Found TeamId {cachedValue} for group {groupId} in cache");

                return cachedValue;
            }

            try
            {
                foundTeam = (await graphClient.Groups[groupId].Team.GetAsync())?.Id;
                redisDB.StringSet($"TeamId for: {groupId}", foundTeam);

                if(debug)
                    log?.LogInformation($"GetTeamFromGroup: Found TeamId {foundTeam} for group {groupId}");
            }
            catch (Exception ex)
            {
                log?.LogError("GetTeamFromGroup: " + ex.ToString());
            }

            return foundTeam;
        }

        public async Task<Team?> CreateTeamFromGroup(string groupId, bool debug)
        {
            Team? createdTeam = null;
            string? createdTeamId = "";

            if(graphClient == null || redisDB == null)
            {
                return createdTeam;
            }

            try
            {
                createdTeamId = await this.GetTeamFromGroup(groupId, debug);
            }
            catch (Exception ex)
            {
                log?.LogError("CreateTeamFromGroup: " + ex.ToString());
            }

            if (string.IsNullOrEmpty(createdTeamId))
            {
                if(debug)
                    log?.LogInformation("CreateTeamFromGroup: Creating team for group " + groupId);

                try
                {
                    Team teamSettings = new Team()
                    {
                        MemberSettings = new TeamMemberSettings()
                        {
                            AllowCreatePrivateChannels = true,
                            AllowCreateUpdateChannels = true
                        },
                        MessagingSettings = new TeamMessagingSettings()
                        {
                            AllowUserEditMessages = true,
                            AllowUserDeleteMessages = true
                        },
                        FunSettings = new TeamFunSettings()
                        {
                            AllowGiphy = true,
                            GiphyContentRating = GiphyRatingType.Strict
                        }
                    };

                    //create a team from newly created group
                    createdTeam = await graphClient.Groups[groupId].Team.PutAsync(teamSettings);
                    redisDB.StringSet($"TeamId for: {groupId}", createdTeam?.Id);

                    if(debug)
                        log?.LogInformation("CreateTeamFromGroup: Waiting 60s for team to be created");

                    //wait for team to be created
                    Thread.Sleep(60000);
                }
                catch (Exception ex)
                {
                    log?.LogError("CreateTeamFromGroup: " + ex.Message);
                }
            }

            return createdTeam;
        }

        public async Task<string?> AddTeamApp(string teamId, string appId, bool debug)
        {
            string? returnValue = "";

            if(graphClient == null || !string.IsNullOrEmpty(teamId))
            {
                return returnValue;
            }

            try
            {
                if(string.IsNullOrEmpty(await this.IsTeamInstalledApp(teamId, "", appId, debug)))
                {
                    if(debug)
                        log?.LogInformation("AddTeamApp: Add app to team " + teamId);

                    var teamsAppInstallation = new TeamsAppInstallation
                    {
                        AdditionalData = new Dictionary<string, object>()
                    {
                        {"teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + appId}
                    }
                    };

                    var installation = await graphClient.Teams[teamId].InstalledApps.PostAsync(teamsAppInstallation);
                    returnValue = installation?.TeamsApp?.Id;
                }
            }
            catch (Exception ex)
            {
                log?.LogError("AddTeamApp: " + ex.Message);
            }

            return returnValue;
        }

        public async Task<string?> IsTeamInstalledApp(string teamId, string appName, string appId, bool debug)
        {
            string? returnValue = "";

            if(redisDB == null || graphClient == null)
            {
                return returnValue;
            }

            if (!string.IsNullOrEmpty(appName))
            {
                RedisValue cachedValue = redisDB.StringGet($"App: {appName} for: {teamId}");

                if(cachedValue.HasValue && !cachedValue.IsNullOrEmpty)
                {
                    if (debug)
                        log?.LogInformation($"IsTeamInstalledApp: Found app {appName} for team {teamId} in cache");

                    return cachedValue;
                }
            }

            if (!string.IsNullOrEmpty(appId))
            {
                RedisValue cachedValue = redisDB.StringGet($"App: {appId} for: {teamId}");

                if (cachedValue.HasValue && !cachedValue.IsNullOrEmpty)
                {
                    if (debug)
                        log?.LogInformation($"IsTeamInstalledApp: Found appid {appId} for team {teamId} in cache");

                    return cachedValue;
                }
            }

            var apps = await graphClient.Teams[teamId].InstalledApps.GetAsync();

            if (apps?.Value?.Count > 0)
            {
                if (!string.IsNullOrEmpty(appName))
                {
                    returnValue = apps.Value.FirstOrDefault(a => a.TeamsAppDefinition?.DisplayName == appName)?.TeamsApp?.Id;
                    redisDB.StringSet($"App: {appId} for: {teamId}", appId);
                }

                if (!string.IsNullOrEmpty(appId))
                {
                    returnValue = apps.Value.FirstOrDefault(a => a.TeamsAppDefinition?.TeamsAppId == appId)?.TeamsApp?.Id;
                    redisDB.StringSet($"App: {appId} for: {teamId}", appId);
                }
            }

            return returnValue;
        }

        public async Task<TeamsApp?> GetTeamApp(string appName, string appId)
        {
            TeamsApp? returnValue = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            var apps = await graphClient.AppCatalogs.TeamsApps.GetAsync();

            if(apps?.Value?.Count > 0)
            {
                if(!string.IsNullOrEmpty(appName))
                {
                    returnValue = apps.Value.FirstOrDefault(a => a.DisplayName == appName);
                }

                if (!string.IsNullOrEmpty(appId))
                {
                    returnValue = apps.Value.FirstOrDefault(a => a.Id == appId);
                }
            }

            return returnValue;
        }

        public async Task<bool> TabExistsById(string teamId, string channelId, string tabId)
        {
            bool returnValue = false;

            if (graphClient == null)
            {
                return false;
            }

            if (LookupCacheList($"Tabids for team: {teamId} and channel: {channelId}", tabId))
            {
                returnValue = true;
            }
            else
            {
                try
                {
                    var tabs = await graphClient.Teams[teamId].Channels[channelId].Tabs.GetAsync();

                    List<string> tabsCache = new List<string>();

                    if (tabs?.Value?.Count > 0)
                    {
                        var tabsList = tabs.Value.ToList();
                        var tabsSelect = tabsList.Select(i => i.Id).ToList();

                        if(tabsSelect != null)
                        {
                            AddCacheList($"Tabids for team: {teamId} and channel: {channelId}", tabsSelect);
                            returnValue = tabs.Value.Any(t => t.Id == tabId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.Message);
                }
            }

            return returnValue;
        }

        public async Task<bool> TabExists(string teamId, string channelId, string tabName, bool debug)
        {
            bool returnValue = false;

            if(graphClient == null)
            {
                return false;
            }

            if(LookupCacheList($"Tabnames for team: {teamId} and channel: {channelId}", tabName))
            {
                returnValue = true;
            }
            else
            {
                try
                {
                    var tabs = await graphClient.Teams[teamId].Channels[channelId].Tabs.GetAsync();

                    if (tabs?.Value?.Count > 0)
                    {
                        AddCacheList($"Tabnames for team: {teamId} and channel: {channelId}", tabs.Value.Select(i => i.DisplayName).ToList());
                        returnValue = tabs.Value.Any(t => t.DisplayName == tabName);
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("TabExists: " + ex.Message);
                }
            }

            return returnValue;
        }

        public async Task<dynamic?> GetTab(string teamId, string channelId, string tabName, bool debug)
        {
            dynamic? returnValue = null;

            if (graphClient == null)
            {
                return returnValue;
            }

            var tab = LookupCacheList($"Tabs for team: {teamId} and channel: {channelId}", item => item.name == tabName);

            if(tab != null)
            {
                returnValue = new { id = tab.id, name = tab.name };
            }

            try
            {
                var tabs = await graphClient.Teams[teamId].Channels[channelId].Tabs.GetAsync();

                if (tabs != null && tabs.Value != null && tabs.Value.Any(t => t.DisplayName == tabName))
                {
                    AddCacheList($"Tabs for team: {teamId} and channel: {channelId}", tabs.Value.Select(i => new { id = i.Id, name = i.DisplayName }).ToList<dynamic?>());
                    returnValue = tabs?.Value?.FirstOrDefault(t => t.DisplayName == tabName)?.Id;
                }
            }
            catch (Exception ex)
            {
                log?.LogError("GetTab: " + ex.Message);
            }

            return returnValue;
        }

        public async Task RemoveTab(string teamId, string channelId, string TabId, bool debug)
        {
            if (graphClient != null && await TabExists(teamId, channelId, TabId, debug))
            {
                try
                {
                    await graphClient.Teams[teamId].Channels[channelId].Tabs[TabId].DeleteAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("RemoveTab: " + ex.Message);
                }
            }
        }

        public async Task<bool> AddChannelWebApp(string teamId, string channelId, string tabName, string contentUrl, string webUrl, bool debug)
        {
            bool returnValue = false;

            if (!(await TabExists(teamId, channelId, tabName, debug)))
            {
                try
                {
                    if(debug)
                        log?.LogInformation("AddChannelWebApp: Adding website tab");

                    TeamsTab infotab = new TeamsTab()
                    {
                        DisplayName = tabName,
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web" }
                        },
                        Configuration = new TeamsTabConfiguration()
                        {
                            WebsiteUrl = webUrl,
                            ContentUrl = contentUrl
                        }
                    };

                    if (graphClient != null)
                    {
                        var tab = await graphClient.Teams[teamId].Channels[channelId].Tabs.PostAsync(infotab);

                        returnValue = true;
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("AddChannelWebApp: " + ex.Message);
                }
            }

            return returnValue;
        }

        public async Task<bool> AddChannelApp(string teamId, string appId, string channelId, string tabName, string entityId, string contentUrl, string webUrl, string removeUrl, bool debug)
        {
            bool returnValue = false;

            if(!(await TabExists(teamId, channelId, tabName, debug)))
            {
                try
                {
                    if(debug)
                        log?.LogInformation("AddChannelApp: Adding app tab");

                    TeamsTab infotab = new TeamsTab()
                    {
                        DisplayName = tabName,
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + appId }
                        },
                        Configuration = new TeamsTabConfiguration()
                        {
                            ContentUrl = contentUrl
                        }
                    };

                    if (entityId != null)
                    {
                        infotab.Configuration.EntityId = entityId;
                    }

                    if (webUrl != null)
                    {
                        infotab.WebUrl = webUrl;
                    }

                    if(graphClient != null)
                    {
                        var tab = await graphClient.Teams[teamId].Channels[channelId].Tabs.PostAsync(infotab);

                        if(tab != null)
                        {
                            returnValue = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("AddChannelApp: " + ex.Message);
                }
            }

            return returnValue;
        }

        public async Task<string?> FindChannel(string teamId, string channelName, bool debug)
        {
            string? returnValue = null;
            
            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                if(debug)
                    log?.LogInformation("FindChannel: Find channel " + channelName + " in team " + teamId);

                var channels = await graphClient.Teams[teamId].Channels.GetAsync();

                if(channels?.Value?.Count() > 0)
                {
                    returnValue = channels.Value.FirstOrDefault(c => c.DisplayName == channelName)?.Id;

                    if(returnValue != null && debug)
                    {
                        log?.LogInformation("FindChannel: Channel " + channelName + " found in team " + teamId);
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError("FindChannel: " + ex.Message);
            }

            return returnValue;
        }

        public async Task<string?> AddChannel(string teamId, string channelName, string channelDescription, ChannelMembershipType type, bool debug)
        {
            string? returnValue = "";

            if(graphClient == null)
            {
                return returnValue;
            }

            if(string.IsNullOrEmpty(await FindChannel(teamId, channelName, debug)))
            {
                try
                {
                    if(debug)
                        log?.LogInformation("AddChannel: Adding " + channelName + " to team " + teamId);

                    var channel = new Channel
                    {
                        DisplayName = channelName,
                        Description = channelDescription,
                        MembershipType = type
                    };

                    Channel? createdChannel = await graphClient.Teams[teamId].Channels.PostAsync(channel);
                    
                    if(createdChannel != null)
                    {
                        returnValue = createdChannel.Id;
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("AddChannel: " + ex.Message);
                }
            }
            else
            {
                log?.LogInformation("AddChannel: " + channelName + " already existed in team " + teamId);
            }

            return returnValue;
        }

        public async Task CreatePlannerTabInChannelAsync(string teamId, string tabName, string channelId, string planId)
        {
            if(settings == null || graphClient == null)
            {
                return;
            }

            var tab = new TeamsTab
            {
                DisplayName = tabName,
                AdditionalData = new Dictionary<string, object>()
                {
                    { "teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web" }
                },
                Configuration = new TeamsTabConfiguration
                {
                    ContentUrl = $"https://tasks.office.com/{settings.TenantID}/en-US/Home/PlannerFrame?page=7&planId={planId}&auth=true",
                    WebsiteUrl = $"https://tasks.office.com/{settings.TenantID}/en-US/Home/PlanViews/{planId}"
                }
            };

            try
            {
                var createdTab = await graphClient.Teams[teamId].Channels[channelId].Tabs.PostAsync(tab);

                if(createdTab != null)
                {
                    Console.WriteLine($"Planner tab created with ID: {createdTab.Id}");
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating Planner tab: {ex.Message}");
            }
        }
        #endregion

        #region Groups
        public async Task<bool> AddGroupOwner(string userEmail, string GroupId, bool debug)
        {
            string? memberToAdd = "";
            bool returnValue = true;

            if (redisDB == null)
            {
                return returnValue;
            }

            var cacheValue = redisDB.StringGet(userEmail);

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null)
            {
                if(!cacheValue.IsNullOrEmpty && cacheValue.HasValue)
                {
                    memberToAdd = cacheValue;
                }
                else {
                    try
                    {
                        if(debug)
                            log?.LogInformation($"AddGroupOwner: Trying to find member {userEmail}");

                        memberToAdd = (await graphClient.Users[userEmail].GetAsync())?.Id;
                        redisDB.StringSet(userEmail, memberToAdd);
                    }
                    catch (Exception ex)
                    {
                        log?.LogError($"AddGroupOwner: Unable to find member {userEmail} with error {ex.ToString()}");
                    }
                }

                if (!string.IsNullOrEmpty(memberToAdd))
                {
                    var directoryObject = new ReferenceCreate
                    {
                        OdataId = "https://graph.microsoft.com/v1.0/directoryObjects/" + memberToAdd
                    };

                    bool ownerExists = false;

                    try
                    {
                        var owners = await graphClient.Groups[GroupId].Owners.GetAsync();

                        if (owners?.Value?.Count > 0)
                        {
                            foreach (var owner in owners.Value)
                            {
                                if (owner.Id == memberToAdd)
                                {
                                    ownerExists = true;
                                }
                            }

                            if (!ownerExists)
                            {
                                if(debug)
                                    log?.LogInformation($"AddGroupOwner: Adding owner {userEmail}: {memberToAdd} to group");

                                await graphClient.Groups[GroupId].Owners.Ref.PostAsync(directoryObject);
                                ownerExists = true;
                            }
                        }
                        else
                        {
                            if(debug)
                                log?.LogInformation($"AddGroupOwner: Adding owner {userEmail}: {memberToAdd} to group");

                            await graphClient.Groups[GroupId].Owners.Ref.PostAsync(directoryObject);
                            ownerExists = true;
                        }

                        returnValue = true;
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("AddGroupOwner: " + ex.Message);
                        returnValue = false;
                    }
                }
                else
                {
                    returnValue = false;
                }
            }

            return returnValue;
        }

        public async Task<bool> AddGroupMember(string userEmail, string GroupId, bool debug)
        {
            string? memberToAdd = "";
            bool returnValue = true;

            if(redisDB == null)
            {
                return returnValue;
            }

            var cacheValue = redisDB.StringGet(userEmail);

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null)
            {
                if (!cacheValue.IsNullOrEmpty && cacheValue.HasValue)
                {
                    memberToAdd = cacheValue;
                }
                else
                {
                    try
                    {
                        if(debug)
                            log?.LogInformation($"AddGroupMember: Trying to find member {userEmail}");

                        memberToAdd = (await graphClient.Users[userEmail].GetAsync())?.Id;
                        redisDB.StringSet(userEmail, memberToAdd);
                    }
                    catch (Exception ex)
                    {
                        log?.LogError($"AddGroupMember: Unable to find member {userEmail} with error {ex.ToString()}");
                    }
                }

                if (!string.IsNullOrEmpty(memberToAdd))
                {
                    var directoryObject = new ReferenceCreate
                    {
                        OdataId = "https://graph.microsoft.com/v1.0/directoryObjects/" + memberToAdd
                    };

                    bool memberExists = false;

                    try
                    {
                        var members = await graphClient.Groups[GroupId].Members.GetAsync();

                        if(members?.Value?.Count > 0)
                        {
                            foreach(var member in members.Value)
                            {
                                if(member.Id == memberToAdd)
                                {
                                    memberExists = true;
                                }
                            }

                            if(!memberExists)
                            {
                                if(debug)
                                    log?.LogInformation($"AddGroupMember: Adding owner {userEmail}: {memberToAdd} to group");

                                await graphClient.Groups[GroupId].Members.Ref.PostAsync(directoryObject);
                                memberExists = true;
                            }
                        }
                        else
                        {
                            if(debug)
                                log?.LogInformation($"AddGroupMember: Adding owner {userEmail}: {memberToAdd} to group");

                            await graphClient.Groups[GroupId].Members.Ref.PostAsync(directoryObject);
                            memberExists = true;
                        }

                        returnValue = true;
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("AddGroupMember: " + ex.Message);
                        returnValue = false;
                    }
                }
                else
                {
                    returnValue = false;
                }
            }

            return returnValue;
        }

        public async Task<Group?> CreateGroup(string description, string mailNickname, List<string> owners, bool debug)
        {
            Group? createdGroup = default(Group);

            if (!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname) && graphClient != null)
            {
                if (debug)
                {
                    log?.LogInformation("CreateGroup: Creating group " + description);
                    log?.LogInformation("CreateGroup: With owners " + String.Join(",", owners));
                }

                try
                {
                    Group? createGroupBody = GetCreateGroupBody(description, mailNickname, owners);

                    if(createGroupBody != null)
                    {
                        createdGroup = await graphClient.Groups.PostAsync(createGroupBody);
                    }
                }
                catch (Exception ex)
                {
                    if(debug)
                        log?.LogInformation("CreateGroup: Error creating group.");

                    log?.LogError("CreateGroup: " + ex.Message);
                }

                if(debug)
                    log?.LogInformation("CreateGroup: Waiting 60s for group to be created.");

                //wait for group to be created
                Thread.Sleep(90000);
            }

            return createdGroup;
        }

        public async Task<Group?> CreateGroupNoWait(string description, string mailNickname, List<string> owners, bool debug)
        {
            Group? createdGroup = default(Group);

            if (!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname) && graphClient != null)
            {
                if (debug)
                {
                    log?.LogInformation("CreateGroupNoWait: Creating group " + description);
                    log?.LogInformation("CreateGroupNoWait: With owners " + String.Join(",", owners));
                }

                try
                {
                    Group? createGroupBody = GetCreateGroupBody(description, mailNickname, owners);

                    if(createGroupBody != null)
                    {
                        createdGroup = await graphClient.Groups.PostAsync(createGroupBody);
                    }
                }
                catch (Exception ex)
                {
                    if(debug)
                        log?.LogInformation("CreateGroupNoWait: Error creating group.");
                    
                    log?.LogError("CreateGroupNoWait: " + ex.Message);
                }
            }

            return createdGroup;
        }

        public Group? GetCreateGroupBody(string description, string mailNickname, List<string> owners)
        {
            Group? createGroupBody = default(Group);

            if (!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname) && owners.Count > 0)
            {
                createGroupBody = new Group()
                {
                    Description = description,
                    DisplayName = description,
                    GroupTypes = new List<string>() { "Unified" },
                    MailEnabled = true,
                    MailNickname = mailNickname,
                    SecurityEnabled = false,
                    AdditionalData = new Dictionary<string, object>()
                    {
                        { "owners@odata.bind", owners.Distinct().ToArray() }
                    }
                };

            }
            else if(!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname))
            {
                createGroupBody = new Group()
                {
                    Description = description,
                    DisplayName = description,
                    GroupTypes = new List<string>() { "Unified" },
                    MailEnabled = true,
                    MailNickname = mailNickname,
                    SecurityEnabled = false
                };
            }

            return createGroupBody;
        }

        public async Task<bool> GroupExists(string mailNickname)
        {
            if(graphClient == null)
            {
                return false;
            }

            var existingGroup = await graphClient.Groups.GetAsync(config =>
                {
                    config.QueryParameters.Filter = "mailNickname eq '" + mailNickname + "'";
                    config.QueryParameters.Select = new string[] { "id, displayName" };
                });

            if (existingGroup?.Value?.Count <= 0)
            {
                return false;
            }

            return true;
        }

        public async Task<FindGroupResult> GetGroupById(string id, bool debug)
        {
            FindGroupResult returnValue = new FindGroupResult();
            returnValue.Success = false;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                string? group = (await graphClient.Groups[id].GetAsync())?.Id;

                if (group != null)
                {
                    if(debug)
                        log?.LogInformation($"GetGroupById: Found group with id {group}");

                    returnValue.Success = true;
                    returnValue.Count = 1;
                    returnValue.group = group;
                }
            }
            catch (Exception ex)
            {
                log?.LogInformation($"GetGroupById: Group not found. {ex}");
            }

            return returnValue;
        }

        public async Task<FindGroupResult> FindGroupByName(string mailNickname, bool withRetry)
        {
            FindGroupResult returnValue = new FindGroupResult();
            returnValue.Success = false;
            int maxcnt = 2;
            int cnt = 0;
            GroupCollectionResponse? foundGroups = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            //try to find the group (might fail if the group has not been created yet)
            try
            {
                foundGroups = await graphClient.Groups.GetAsync(config => {
                    config.QueryParameters.Filter = "mailNickname eq '" + mailNickname + "'";
                });
            }
            catch (Exception)
            {
            }

            while (foundGroups?.Value?.Count <= 0 && withRetry)
            {
                lock (returnValue) {
                    log?.LogInformation($"Group not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");
                    Thread.Sleep(10000);
                    cnt+=1;
                }

                try
                {
                    foundGroups = await graphClient.Groups.GetAsync(config => {
                        config.QueryParameters.Filter = "mailNickname eq '" + mailNickname + "'";
                    });
                }
                catch (Exception)
                {
                }

                if (cnt == maxcnt)
                    break;
            }

            if (foundGroups?.Value?.Count > 1)
            {
                returnValue.Success = true;
                returnValue.Count = foundGroups.Value.Count;
                returnValue.groups = foundGroups.Value.Select(g => g.Id).ToList<string?>();
            }
            else if(foundGroups?.Value?.Count > 0)
            {
                returnValue.Success = true;
                returnValue.Count = 1;
                returnValue.group = foundGroups.Value[0].Id;
            }
            else
            {
                returnValue.Success = false;
            }

            return returnValue;
        }
        #endregion

        #region Drives
        public async Task<Drive?> GetGroupDrive(Group? group, bool debug)
        {
            Drive? groupDrive = null;

            if(graphClient == null || group == null)
            {
                return null;
            }

            try
            {
                if (debug)
                    log?.LogInformation($"Trying to get group drive for {group.Id}");

                groupDrive = await graphClient.Groups[group.Id].Drive.GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError("GetGroupDrive: " + ex.ToString());
            }

            return groupDrive;
        }

        public async Task<string?> GetGroupDrive(string? GroupId, bool debug)
        {
            string? groupDrive = "";

            if (graphClient == null || string.IsNullOrEmpty(GroupId) || redisDB == null)
            {
                return null;
            }

            var cachedValue = redisDB.StringGet($"Drive for group: {GroupId}");

            if(!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                if (debug)
                    log?.LogInformation($"Found group drive for {GroupId} in cache");

                groupDrive = cachedValue;
            }
            else
            {
                try
                {
                    if (debug)
                        log?.LogInformation($"Trying to get group drive for {GroupId}");

                    groupDrive = (await graphClient.Groups[GroupId].Drive.GetAsync())?.Id;
                    redisDB.StringSet($"Drive for group: {GroupId}", groupDrive);
                }
                catch (Exception ex)
                {
                    log?.LogError("GetGroupDrive: " + ex.ToString());
                }
            }

            return groupDrive;
        }

        public async Task<string?> GetGroupDriveUrl(string? GroupId, bool debug)
        {
            string? groupDriveUrl = "";

            if (graphClient == null || string.IsNullOrEmpty(GroupId) || redisDB == null)
            {
                return null;
            }

            var cachedValue = redisDB.StringGet($"DriveUrl for group: {GroupId}");

            if (!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                if (debug)
                    log?.LogInformation($"GetGroupDriveUrl: Found group drive url for {GroupId} in cache");
                groupDriveUrl = cachedValue;
            }
            else
            {
                try
                {
                    groupDriveUrl = (await graphClient.Groups[GroupId].Drive.GetAsync())?.WebUrl;

                    if (debug)
                        log?.LogInformation($"GetGroupDriveUrl: Found group drive url for {GroupId}");

                    redisDB.StringSet($"DriveUrl for group: {GroupId}", groupDriveUrl);
                }
                catch (Exception ex)
                {
                    log?.LogError("GetGroupDriveUrl: " + ex.ToString());
                }
            }

            return groupDriveUrl;
        }

        public async Task<string?> GetGroupSite(string? GroupId, bool debug)
        {
            string? returnValue = "";

            if (graphClient == null || string.IsNullOrEmpty(GroupId) || redisDB == null)
            {
                return null;
            }

            var cachedValue = redisDB.StringGet($"Site for group: {GroupId}");
            
            if(!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                if(debug)
                    log?.LogInformation($"GetGroupSite: Found group id {GroupId} in cache");

                return cachedValue;
            }

            FindGroupResult? findGroup = await GetGroupById(GroupId, debug);

            if(findGroup?.Success == true && findGroup?.group != null)
            {
                var sites = await graphClient.Groups[findGroup.group].Sites.GetAsync();

                if(sites?.Value?.Count > 0)
                {
                    returnValue = sites?.Value[0].Id;
                    redisDB.StringSet($"Site for group: {GroupId}", returnValue);

                    if (debug)
                        log?.LogInformation($"GetGroupSite: Found group id {returnValue}");
                }
            }

            return returnValue;
        }

        //public async Task<Drive?> GetSiteDrive(Site? site)
        //{
        //    Drive? groupDrive = null;

        //    if (graphClient == null || site == null)
        //    {
        //        return null;
        //    }

        //    try
        //    {
        //        groupDrive = await graphClient.Sites[site.Id].Drive.GetAsync();
        //    }
        //    catch (Exception ex)
        //    {
        //        log?.LogError(ex.ToString());
        //    }

        //    return groupDrive;
        //}

        public async Task<string?> GetSiteDrive(string? SiteId, bool debug)
        {
            string? groupDriveId = "";

            if (graphClient == null || string.IsNullOrEmpty(SiteId) || redisDB == null)
            {
                return null;
            }

            var cachedValue = redisDB.StringGet($"Drive for site: {SiteId}");

            if(!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                if (debug)
                    log?.LogInformation($"GetSiteDrive: Found drive for site {SiteId} in cache");

                groupDriveId = cachedValue;
            }

            try
            {
                groupDriveId = (await graphClient.Sites[SiteId].Drive.GetAsync())?.Id;

                if (debug)
                    log?.LogInformation($"GetSiteDrive: Found drive for site {SiteId}");

                redisDB.StringSet($"Drive for site: {SiteId}", groupDriveId);
            }
            catch (Exception ex)
            {
                log?.LogError("GetSiteDrive: " + ex.ToString());
            }

            return groupDriveId;
        }

        public async Task<List<DriveItem>> GetDriveRootItems(Drive? groupDrive, bool debug)
        {
            List<DriveItem> returnValue = new List<DriveItem>();

            if (groupDrive != null && graphClient != null)
            {
                DriveItem? root = null;

                try
                {
                    root = await graphClient.Drives[groupDrive.Id].Root.GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("GetDriveRootItems: " + ex.ToString());
                }

                if (root != null && graphClient != null)
                {
                    DriveItemCollectionResponse? rootChildren = null;

                    try
                    {
                        rootChildren = await graphClient.Drives[groupDrive.Id].Items[root.Id].Children.GetAsync();
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("GetDriveRootItems: " + ex.ToString());
                    }

                    if(rootChildren?.Value?.Count > 0)
                    {
                        returnValue = rootChildren.Value.ToList();
                    }
                }
            }

            return returnValue;
        }

        public async Task<List<DriveItem>> GetDriveRootItems(string? groupDriveId, bool debug)
        {
            List<DriveItem> returnValue = new List<DriveItem>();

            if (!string.IsNullOrEmpty(groupDriveId) && graphClient != null)
            {
                DriveItem? root = null;

                try
                {
                    root = await graphClient.Drives[groupDriveId].Root.GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("GetDriveRootItems: " + ex.ToString());
                }

                if (root != null && graphClient != null)
                {
                    DriveItemCollectionResponse? rootChildren = null;

                    try
                    {
                        rootChildren = await graphClient.Drives[groupDriveId].Items[root.Id].Children.GetAsync();
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("GetDriveRootItems: " + ex.ToString());
                    }

                    if (rootChildren?.Value?.Count > 0)
                    {
                        returnValue = rootChildren.Value.ToList();
                    }
                }
            }

            return returnValue;
        }

        public async Task<List<DriveItem>> GetDriveFolderChildren(string? groupDriveId, string? parentId, bool recursive = false, bool debug = false)
        {
            List<DriveItem> returnValue = new List<DriveItem>();

            if (!string.IsNullOrEmpty(groupDriveId) && graphClient != null && !string.IsNullOrEmpty(parentId))
            {
                var folderChildren = await graphClient.Drives[groupDriveId].Items[parentId].Children.GetAsync();

                if (folderChildren?.Value?.Count > 0)
                {
                    if (recursive)
                    {
                        foreach(var child in folderChildren.Value)
                        {
                            var subchildren = await GetDriveFolderChildren(groupDriveId, child.Id, recursive, debug);

                            if(subchildren?.Count > 0)
                            {
                                child.Children = subchildren;
                            }
                        }
                    }

                    returnValue = folderChildren.Value;
                }
            }

            return returnValue;
        }

        //public async Task<List<DriveItem>> GetDriveFolderChildren(string? groupDriveId, string? parentId, bool recursive = false)
        //{
        //    List<DriveItem> returnValue = new List<DriveItem>();

        //    if (!string.IsNullOrEmpty(groupDriveId) && graphClient != null && !string.IsNullOrEmpty(parentId))
        //    {
        //        var folderChildren = await graphClient.Drives[groupDriveId].Items[parentId].Children.GetAsync();

        //        if (folderChildren?.Value?.Count > 0)
        //        {
        //            if (recursive)
        //            {
        //                foreach (var child in folderChildren.Value)
        //                {
        //                    var subchildren = await GetDriveFolderChildren(groupDriveId, child.Id, recursive);

        //                    if (subchildren?.Count > 0)
        //                    {
        //                        child.Children = subchildren;
        //                    }
        //                }
        //            }

        //            returnValue = folderChildren.Value;
        //        }
        //    }

        //    return returnValue;
        //}

        public async Task<DriveItem?> FindItem(Drive? groupDrive, string? Path, bool withRetry = false, bool debug = false)
        {
            DriveItem? returnValue = null;

            int maxcnt = 2;
            int cnt = 0;

            if(graphClient == null || groupDrive == null || string.IsNullOrEmpty(Path))
            {
                return null;
            }

            try
            {
                returnValue = await graphClient.Drives[groupDrive.Id].Root.ItemWithPath(Path).GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError("FindItem: " + ex.ToString());
            }

            while (returnValue == null && withRetry)
            {
                if(returnValue != null)
                {
                    lock (returnValue)
                    {
                        if(debug)
                            log?.LogInformation($"FindItem: Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");

                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDrive.Id].Root.ItemWithPath(Path).GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("FindItem: " + ex.ToString());
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DriveItem?> FindItem(string? groupDriveId, string? Path, bool withRetry = false, bool debug = false)
        {
            DriveItem? returnValue = null;

            int maxcnt = 2;
            int cnt = 0;

            if (graphClient == null || string.IsNullOrEmpty(groupDriveId) || string.IsNullOrEmpty(Path))
            {
                return null;
            }

            try
            {
                returnValue = await graphClient.Drives[groupDriveId].Root.ItemWithPath(Path).GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError("FindItem: " + ex.ToString());
            }

            while (returnValue == null && withRetry)
            {
                if (returnValue != null)
                {
                    lock (returnValue)
                    {
                        if(debug)
                            log?.LogInformation($"FindItem: Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");

                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDriveId].Root.ItemWithPath(Path).GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("FindItem: " + ex.ToString());
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DriveItem?> FindItem(Drive? groupDrive, string? parentId, string? Path, bool withRetry, bool debug)
        {
            DriveItem? returnValue = null;

            if(graphClient == null || groupDrive == null || string.IsNullOrEmpty(parentId) || string.IsNullOrEmpty(Path)) 
            { 
                return null; 
            }

            try
            {
                returnValue = await graphClient.Drives[groupDrive.Id].Items[parentId].ItemWithPath(Path).GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError("FindItem: " + ex.ToString());
            }

            int maxcnt = 2;
            int cnt = 0;

            while (returnValue == null && withRetry)
            {
                if(returnValue != null)
                {
                    lock (returnValue)
                    {
                        if(debug)
                            log?.LogInformation($"FindItem: Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");

                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDrive.Id].Items[parentId].ItemWithPath(Path).GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("FindItem: " + ex.ToString());
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DriveItem?> FindItem(string? groupDriveId, string? parentId, string? Path, bool withRetry, bool debug)
        {
            DriveItem? returnValue = null;

            if (graphClient == null || string.IsNullOrEmpty(groupDriveId) || string.IsNullOrEmpty(parentId) || string.IsNullOrEmpty(Path))
            {
                return null;
            }

            try
            {
                returnValue = await graphClient.Drives[groupDriveId].Items[parentId].ItemWithPath(Path).GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError("FindItem: " + ex.ToString());
            }

            int maxcnt = 2;
            int cnt = 0;

            while (returnValue == null && withRetry)
            {
                if (returnValue != null)
                {
                    lock (returnValue)
                    {
                        if(debug)
                            log?.LogInformation($"FindItem: Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");

                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDriveId].Items[parentId].ItemWithPath(Path).GetAsync();
                }
                catch (Exception ex)
                {
                    log?.LogError("FindItem: " + ex.ToString());
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DownloadFileResult> DownloadFile(string? GroupID, string? FolderID, string? FileName, bool debug)
        {
            DownloadFileResult returnValue = new DownloadFileResult();
            Stream orderFileStream = Stream.Null;

            if(graphClient == null || string.IsNullOrEmpty(GroupID) || string.IsNullOrEmpty(FolderID) || string.IsNullOrEmpty(FileName))
            {
                return returnValue;
            }

            try
            {
                if(debug)
                    log?.LogInformation($"DownloadFile: Trying to find group drive for file {FileName}");

                string? groupDriveId = await GetGroupDrive(GroupID, debug);

                if(!string.IsNullOrEmpty(groupDriveId))
                {
                    if(debug)
                        log?.LogInformation($"DownloadFile: Found group drive for file {FileName}");

                    //download order file content
                    var stream = await graphClient.Drives[groupDriveId].Items[FolderID].ItemWithPath(FileName).Content.GetAsync();

                    if(stream != null && stream != Stream.Null)
                    {
                        returnValue.Contents = new MemoryStream();
                        await stream.CopyToAsync(returnValue.Contents);

                        if (debug)
                            log?.LogInformation($"DownloadFile: Downloaded file {FileName} with size {returnValue.Contents?.Length} byte");

                        returnValue.Success = true;
                    }
                }
            }
            catch (Exception ex)
            {
                returnValue.Success = false;
                log?.LogError("DownloadFile: " + ex.ToString());
            }

            return returnValue;
        }

        public async Task<DownloadFileResult> DownloadFile(Group? Group, DriveItem? Folder, string? Path, bool debug)
        {
            DownloadFileResult returnValue = new DownloadFileResult();
            Stream orderFileStream = Stream.Null;

            if (graphClient == null || Group == null || Folder == null || string.IsNullOrEmpty(Path))
            {
                return returnValue;
            }

            try
            {
                if(debug)
                    log?.LogInformation($"DownloadFile: Trying to find group drive for file {Path}");

                Drive? groupDrive = await GetGroupDrive(Group, debug);

                if(groupDrive != null)
                {
                    if(debug)
                        log?.LogInformation($"DownloadFile: Found group drive for file {Path}");

                    //download order file content
                    returnValue.Contents = new MemoryStream();
                    var source = await graphClient.Drives[groupDrive.Id].Items[Folder.Id].ItemWithPath(Path).Content.GetAsync();
                    
                    if(source != null && source != Stream.Null)
                    {
                        await source.CopyToAsync(returnValue.Contents);

                        if (debug)
                            log?.LogInformation($"DownloadFile: Downloaded file {Path} with size {returnValue.Contents.Length} byte");

                        returnValue.Success = true;
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError("DownloadFile: " + ex.ToString());
                returnValue.Success = false;
            }

            return returnValue;
        }

        public async Task<bool> UploadFile(string? GroupID, string? FolderID, string? FileName, MemoryStream FileContents, bool debug)
        {
            bool returnValue = false;

            if (graphClient == null || FileContents == null || FileContents == Stream.Null || string.IsNullOrEmpty(GroupID) || string.IsNullOrEmpty(FolderID) || string.IsNullOrEmpty(FileName) || FileContents.Length <= 0)
            {
                return returnValue;
            }

            try
            {
                if(debug)
                    log?.LogInformation($"UploadFile: Trying to find group drive for file {FileName}");

                string? groupDriveId = await GetGroupDrive(GroupID, debug);

                if(!string.IsNullOrEmpty(groupDriveId))
                {
                    if(debug)
                        log?.LogInformation($"UploadFile: Found group drive for file {FileName}. Creating upload request for stream with size {FileContents.Length} byte");

                    CreateUploadSessionPostRequestBody uploadRequest = new CreateUploadSessionPostRequestBody
                    {
                        Item = new DriveItemUploadableProperties
                        {
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "@microsoft.graph.conflictBehavior", "replace" }
                            }
                        }
                    };

                    if(debug)
                        log?.LogInformation($"UploadFile: Creating upload session");

                    var fileUploadSession = await graphClient.Drives[groupDriveId].Items[FolderID].ItemWithPath(FileName).CreateUploadSession.PostAsync(uploadRequest);

                    if (fileUploadSession != null)
                    {
                        if(debug)
                            log?.LogInformation($"UploadFile: Created upload session");

                        var fileUploadTask = new LargeFileUploadTask<DriveItem>(fileUploadSession, FileContents, maxUploadChunkSize, graphClient.RequestAdapter);

                        var totalLength = FileContents.Length;
                        // Create a callback that is invoked after each slice is uploaded
                        IProgress<long> progress = new Progress<long>(prog => {
                            if(debug)
                                log?.LogInformation($"UploadFile: Uploaded {prog} bytes of {totalLength} bytes");
                        });

                        // Upload the file
                        var uploadResult = await fileUploadTask.UploadAsync(progress);
                        string info = uploadResult.UploadSucceeded ?
                                $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                                "Upload failed";

                        if (debug)
                            log?.LogInformation($"UploadFile: {info}");

                        returnValue = true;
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError($"UploadFile: Error uploading with error {ex.ToString()}");
                returnValue = false;
            }

            return returnValue;
        }

        public async Task<bool> UploadFile(Group? Group, DriveItem? Folder, string? Path, MemoryStream FileContents, bool debug)
        {
            bool returnValue = false;

            if(graphClient == null || Group == null || Folder == null || string.IsNullOrEmpty(Path) || FileContents == null || FileContents == Stream.Null || FileContents.Length <= 0)
            {
                return returnValue;
            }

            try
            {
                string? groupDriveId = await GetGroupDrive(Group.Id, debug);

                if (!string.IsNullOrEmpty(groupDriveId))
                {
                    CreateUploadSessionPostRequestBody uploadRequest = new CreateUploadSessionPostRequestBody
                    {
                        Item = new DriveItemUploadableProperties
                        {
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "@microsoft.graph.conflictBehavior", "replace" }
                            }
                        }
                    };
                    var fileUploadSession = await graphClient.Drives[groupDriveId].Items[Folder.Id].ItemWithPath(Path).CreateUploadSession.PostAsync(uploadRequest);

                    if (fileUploadSession != null)
                    {
                        var fileUploadTask = new LargeFileUploadTask<DriveItem>(fileUploadSession, FileContents, maxUploadChunkSize, graphClient.RequestAdapter);

                        var totalLength = FileContents.Length;
                        // Create a callback that is invoked after each slice is uploaded
                        IProgress<long> progress = new Progress<long>(prog => {
                            if(debug)
                                log?.LogInformation($"UploadFile: Uploaded {prog} bytes of {totalLength} bytes");
                        });

                        // Upload the file
                        var uploadResult = await fileUploadTask.UploadAsync(progress);
                        string info = uploadResult.UploadSucceeded ?
                            $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                            "Upload failed";

                        if(debug)
                            log?.LogInformation($"UploadFile: {info}");

                        returnValue = true;
                    }
                }
            }
            catch (Exception ex)
            {
                returnValue = false;
                log?.LogError("UploadFile: " + ex.ToString());
            }

            return returnValue;
        }

        public async Task<bool> CopyFile(CopyItem? source, CopyItem? destination, bool debug)
        {
            bool returnValue = false;

            if(source != null && destination != null && !string.IsNullOrEmpty(source.GroupId) && !string.IsNullOrEmpty(source.FolderId) && !string.IsNullOrEmpty(source.Path))
            {
                //download the file
                if(debug)
                    log?.LogInformation("CopyFile: Download file " + source.Path);

                DownloadFileResult downloadFile = await this.DownloadFile(source.GroupId, source.FolderId, source.Path, debug);

                if (downloadFile.Success && !string.IsNullOrEmpty(destination.GroupId) && !string.IsNullOrEmpty(destination.FolderId) && !string.IsNullOrEmpty(destination.Path))
                {
                    if(debug)
                        log?.LogInformation("CopyFile: Upload file " + destination.Path);

                    if (await this.UploadFile(destination.GroupId, destination.FolderId, destination.Path, downloadFile.Contents, debug))
                    {
                        returnValue = true;
                    }
                }
            }

            return returnValue;
        }

        public async Task<bool> MoveFile(CopyItem? source, CopyItem? destination, bool debug)
        {
            bool returnValue = false;

            if (graphClient != null && source != null && destination != null && !string.IsNullOrEmpty(source.GroupId) && !string.IsNullOrEmpty(source.FolderId) && !string.IsNullOrEmpty(source.Path))
            {
                if(debug)
                    log?.LogInformation($"MoveFile: Downloading file {source.Path}");

                //download the file
                DownloadFileResult downloadFile = await this.DownloadFile(source.GroupId, source.FolderId, source.Path, debug);

                if (downloadFile.Success && !string.IsNullOrEmpty(destination.GroupId) && !string.IsNullOrEmpty(destination.FolderId) && !string.IsNullOrEmpty(destination.Path))
                {
                    if(debug)
                        log?.LogInformation($"MoveFile: Uploading file {destination.Path}");

                    if (await this.UploadFile(destination.GroupId, destination.FolderId, destination.Path, downloadFile.Contents, debug))
                    {
                        try
                        {
                            string? groupDriveId = await GetGroupDrive(source.GroupId, debug);

                            if(!string.IsNullOrEmpty(groupDriveId))
                            {
                                if(debug)
                                    log?.LogInformation($"MoveFile: Deleting file {source.FileId}");

                                await graphClient.Drives[groupDriveId].Items[source.FileId].DeleteAsync();
                                returnValue = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            log?.LogError("MoveFile: " + ex.ToString());
                        }
                    }
                }
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CopyFolder(string? GroupId, string? ParentId, DriveItem? Folder, bool recursive = false, bool includeFiles = false, bool debug = false)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;

            if(string.IsNullOrEmpty(GroupId) || string.IsNullOrEmpty(ParentId) || Folder == null || string.IsNullOrEmpty(Folder.Name))
            {
                return returnValue;
            }

            DriveItem? createdFolder = null;
            CreateFolderResult result = await this.CreateFolder(GroupId, ParentId, Folder.Name, debug);

            if (result.Success && Folder.Children != null)
            {
                if(debug)
                    log?.LogInformation("CopyFolder: Created " + Folder.Name + " folder.");

                createdFolder = result.folder;

                if (recursive && createdFolder != null)
                {
                    createdFolder.Children = new List<DriveItem>();

                    foreach (var childFolder in Folder.Children)
                    {
                        if (childFolder.Folder == null)
                            continue;

                        var createdChild = await this.CopyFolder(GroupId, createdFolder.Id, childFolder, recursive, includeFiles, debug);

                        if (createdChild?.Success == true && createdChild.folder != null)
                        {
                            createdFolder.Children.Add(createdChild.folder);
                        }
                    }
                }

                if (includeFiles && createdFolder != null)
                {
                    foreach (var childFile in Folder.Children)
                    {
                        if (childFile.Folder != null)
                            continue;

                        CopyItem source = new CopyItem() { GroupId = GroupId, FolderId = Folder.Id ?? "", Path = childFile.Name ?? "" };
                        CopyItem destination = new CopyItem() { GroupId = GroupId, FolderId = createdFolder.Id ?? "", Path = childFile.Name ?? "" };
                        await this.CopyFile(source, destination, debug);

                        if (debug)
                            log?.LogInformation("CopyFolder: Copied " + childFile.Name);
                    }
                }

                returnValue.folder = createdFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CopyFolder(string GroupId, DriveItem Folder, bool recursive, bool? includeFiles, bool debug)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;

            if(debug)
                log?.LogInformation("CopyFolder: Creating " + Folder.Name + " folder.");

            DriveItem? createdFolder = null;
            CreateFolderResult result = new CreateFolderResult();
            result.Success = false;

            if (!string.IsNullOrEmpty(Folder.Name))
            {
                result = this.CreateFolder(GroupId, Folder.Name, debug).Result;
            }

            if (result.Success && Folder.Children != null && result.folder != null)
            {
                createdFolder = result.folder;

                if (recursive)
                {
                    createdFolder.Children = new List<DriveItem>();

                    foreach (var childFolder in Folder.Children)
                    {
                        if (childFolder.Folder == null)
                            continue;

                        var createdChild = await this.CopyFolder(GroupId, childFolder, recursive, includeFiles, debug);

                        if (createdChild.Success && createdChild.folder != null)
                        {
                            createdFolder.Children.Add(createdChild.folder);
                        }
                    }
                }

                if (includeFiles.HasValue && includeFiles.Value == true)
                {
                    foreach (var childFile in Folder.Children.Where(c => c.Folder == null))
                    {
                        if (childFile.Folder != null)
                            continue;

                        CopyItem source = new CopyItem() { GroupId = GroupId, FolderId = Folder.Id ?? "", Path = childFile.Name ?? "" };
                        CopyItem destination = new CopyItem() { GroupId = GroupId, FolderId = createdFolder.Id ?? "", Path = childFile.Name ?? "" };
                        await this.CopyFile(source, destination, debug);
                    }
                }

                returnValue.folder = createdFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CreateFolder(string GroupId, string ParentId, string FolderName, bool debug)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;
            DriveItem? createdFolder = null;

            //first check if folder exists
            if (debug)
                log?.LogInformation("CopyFolder: Check if folder " + FolderName + " exists.");

            var driveId = this.GetGroupDrive(GroupId, debug).Result;
            var existingFolder = this.FindItem(driveId, ParentId, FolderName, false, debug).Result;

            if(existingFolder == null && graphClient != null && !string.IsNullOrEmpty(driveId))
            {
                //if not, create it. fail operation if folder does exist
                try
                {
                    var driveItemFolder = new DriveItem
                    {
                        Name = FolderName,
                        Folder = new Folder
                        {
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"@microsoft.graph.conflictBehavior", "fail"}
                        }
                    };

                    createdFolder = await graphClient.Drives[driveId].Items[ParentId].Children.PostAsync(driveItemFolder);

                    if (createdFolder != null)
                    {
                        returnValue.folder = createdFolder;
                        returnValue.Success = true;
                        returnValue.Existed = false;

                        if(debug)
                            log?.LogInformation("CopyFolder: Created " + FolderName + " folder.");
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CopyFolder: " + ex.ToString());
                }
            }
            else
            {
                if (debug)
                    log?.LogInformation("CopyFolder: Folder " + FolderName + " already existed.");

                returnValue.Existed = true;
                returnValue.folder = existingFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CreateFolder(string GroupId, string FolderName, bool debug)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;
            DriveItem? createdFolder = null;

            //first check if folder exists
            if (debug)
                log?.LogInformation("CreateFolder: Check if folder " + FolderName + " exists.");

            var driveId = this.GetGroupDrive(GroupId, debug).Result;
            var existingFolder = this.FindItem(driveId, FolderName, false, debug).Result;

            if (existingFolder == null && !string.IsNullOrEmpty(driveId) && graphClient != null)
            {
                //if not, create it. fail operation if folder does exist
                try
                {
                    var driveItemFolder = new DriveItem
                    {
                        Name = FolderName,
                        Folder = new Folder
                        {
                        },
                        AdditionalData = new Dictionary<string, object>()
                    {
                        {"@microsoft.graph.conflictBehavior", "fail"}
                    }
                    };

                    var rootItem = await graphClient.Drives[driveId].Root.GetAsync();
                    
                    if(rootItem != null)
                    {
                        createdFolder = await graphClient.Drives[driveId].Items[rootItem.Id].Children.PostAsync(driveItemFolder);

                        if (createdFolder != null)
                        {
                            returnValue.folder = createdFolder;
                            returnValue.Success = true;

                            if(debug)
                                log?.LogInformation("CreateFolder: Created " + FolderName + " folder.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CreateFolder: " + ex.ToString());
                }
            }
            else
            {
                if (debug)
                    log?.LogInformation("CreateFolder: Folder " + FolderName + " already existed.");

                returnValue.folder = existingFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }
        #endregion

        #region List
        public async Task<List<ListItem>> GetListItems(string SiteId, string ListId, string Filter, bool debug)
        {
            List<ListItem> returnValue = new List<ListItem>();

            if(graphClient == null)
            {
                return returnValue;
            }

            var items = await graphClient.Sites[SiteId].Lists[ListId].Items.GetAsync(config =>
            {
                config.QueryParameters.Expand = new string[] { "Fields" };
                config.QueryParameters.Filter = Filter;
            }); 

            while(items?.Value?.Count > 0)
            {
                returnValue = items.Value;
            }

            return returnValue;
        }

        public async Task CreateDriveColumn(string groupId, ColumnDefinition def, bool debug)
        {
            if(graphClient == null)
            {
                return;
            }

            try
            {
                var group = await GetGroupById(groupId, debug);

                if(group.Success)
                {
                    var driveId = await GetGroupDrive(groupId, debug);

                    if(!string.IsNullOrEmpty(driveId))
                    {
                        var list = await graphClient.Drives[driveId].List.GetAsync();
                        var site = await GetGroupSite(groupId, debug);

                        if(list != null && site != null)
                        {
                            if(debug)
                                log?.LogInformation($"CreateDriveColumn: Adding column {def.Description}");

                            var col = await graphClient.Sites[site].Lists[list.Id].Columns.PostAsync(def);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError("CreateDriveColumn: " + ex.ToString());
            }
        }

        public async Task CreateDriveColumn(Site site, List list, ColumnDefinition def, bool debug)
        {
            if(graphClient == null)
            {
                return;
            }

            try
            {
                if(debug)
                    log?.LogInformation($"CreateDriveColumn: Adding column {def.Description}");

                var col = await graphClient.Sites[site.Id].Lists[list.Id].Columns.PostAsync(def);
            }
            catch (Exception ex)
            {
                log?.LogError("CreateDriveColumn: " + ex.ToString());
            }
        }
        #endregion

        #region Plans
        public async Task<PlannerPlan?> CreatePlanAsync(string groupId, string planName, bool debug)
        {
            PlannerPlan? createdPlan = null;

            if (graphClient == null)
            {
                return createdPlan;
            }

            PlannerPlan? newPlan = new PlannerPlan
            {
                Title = planName,
                Container = new PlannerPlanContainer() { ContainerId = groupId, Type = PlannerContainerType.Group }
            };

            try
            {
                createdPlan = await graphClient.Planner.Plans
                    .PostAsync(newPlan);

                if(debug)
                    log?.LogInformation($"CreatePlanAsync: Plan created with ID: {createdPlan?.Id}");
            }
            catch (ServiceException ex)
            {
                log?.LogError($"CreatePlanAsync: Error creating plan wit error {ex.Message}");
            }

            return createdPlan;
        }

        public async Task<IList<PlannerPlan>> GetPlansAsync(string groupId, bool debug)
        {
            List<PlannerPlan> returnValue = new List<PlannerPlan>();

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                var plans = await graphClient.Groups[groupId].Planner.Plans.GetAsync();

                if (plans?.Value?.Count > 0)
                {
                    if(debug)
                        log?.LogInformation("GetPlansAsync: Found " + plans.Value.Count + " plans in group");

                    foreach (PlannerPlan plan in plans.Value)
                    {
                        returnValue.Add(plan);
                    }
                }
            }
            catch (ServiceException ex)
            {
                log?.LogError($"GetPlansAsync: Error retrieving plans with error {ex.Message}");
            }

            return returnValue;
        }

        public async Task<PlannerPlan?> PlanExists(string groupId, string planTitle, bool debug)
        {
            if(debug)
                log?.LogInformation("PlanExists: Trying to find plan " + planTitle + " in group: " + groupId);

            var plans = await GetPlansAsync(groupId, debug);

            if(plans?.Count > 0)
            {
                if(debug)
                    log?.LogInformation("PlanExists: Found " + plans.Count);

                if (plans.Any(p => p.Title == planTitle))
                {
                    return plans.FirstOrDefault(p => p.Title == planTitle);
                }
            }

            return null;
        }

        public async Task<IList<PlannerBucket>> GetBucketsAsync(string planId, bool debug)
        {
            List<PlannerBucket> returnValue = new List<PlannerBucket>();

            if (graphClient == null)
            {
                return returnValue;
            }

            try
            {
                var buckets = await graphClient.Planner.Plans[planId].Buckets
                    .GetAsync();

                if(buckets?.Value?.Count > 0)
                {
                    if (debug)
                        log?.LogInformation($"GetBucketsAsync: Found {buckets.Value.Count} buckets in {planId}");

                    returnValue = buckets.Value;
                }
            }
            catch (ServiceException ex)
            {
                log?.LogError($"GetBucketsAsync: Error retrieving buckets with error {ex.Message}");
            }

            return returnValue;
        }

        public async Task CopyBucketAsync(PlannerBucket sourceBucket, string targetPlanId, bool debug)
        {
            if(graphClient == null)
            {
                return;
            }

            // Create a new bucket in the target plan
            var newBucket = new PlannerBucket
            {
                Name = sourceBucket.Name,
                PlanId = targetPlanId
            };

            PlannerBucket? createdBucket;

            try
            {
                if (debug)
                    log?.LogInformation($"CopyBucketAsync: Creating bucket {sourceBucket.Name} in {targetPlanId}");

                createdBucket = await graphClient.Planner.Buckets.PostAsync(newBucket);
            }
            catch (ServiceException ex)
            {
                log?.LogError($"CopyBucketAsync: Error creating bucket with {ex.Message}");
                return;
            }

            if(createdBucket != null)
            {
                if(debug)
                    log?.LogInformation("CopyBucketAsync: Bucket " + sourceBucket.Name + " created, copying tasks");

                // Retrieve tasks from the source bucket
                var tasks = await graphClient.Planner.Buckets[sourceBucket.Id].Tasks
                    .GetAsync();

                if (tasks?.Value?.Count > 0)
                {
                    // Copy tasks to the new bucket
                    foreach (var task in tasks.Value)
                    {
                        var newTask = new PlannerTask
                        {
                            Title = task.Title,
                            PlanId = targetPlanId,
                            BucketId = createdBucket.Id
                        };

                        try
                        {
                            await graphClient.Planner.Tasks
                                .PostAsync(newTask);
                        }
                        catch (ServiceException ex)
                        {
                            log?.LogError($"CopyBucketAsync: Error creating task with error {ex.Message}");
                        }
                    }
                }
            }
        }

        #endregion

        public bool SearchTabs(dynamic item, string value)
        {
            return (item.id == value || item.name == value);
        }

        public bool LookupCacheList(string key, string lookup)
        {
            if(redisDB == null)
            {
                return false;
            }

            RedisValue cachedValue = redisDB.StringGet(key);

            if(cachedValue.HasValue && !cachedValue.IsNullOrEmpty)
            {
                List<string?>? values = JsonConvert.DeserializeObject<List<string?>?>(cachedValue);
                
                if(values?.Count > 0)
                {
                    return values.Any(v => v == lookup);
                }
            }

            return false;
        }

        public dynamic? LookupCacheList(string key, Func<dynamic, bool> searchFunction)
        {
            if (redisDB == null)
            {
                return null;
            }

            RedisValue cachedValue = redisDB.StringGet(key);

            if (cachedValue.HasValue && !cachedValue.IsNullOrEmpty)
            {
                List<dynamic?>? values = JsonConvert.DeserializeObject<List<dynamic?>?>(cachedValue);

                if (values?.Count > 0)
                {
                    values.Find(item => searchFunction(item));
                }
            }

            return null;
        }

        public bool AddCacheList(string key, string value)
        {
            if (redisDB == null)
            {
                return false;
            }

            RedisValue cachedValue = redisDB.StringGet(key);

            if (!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                string? cval = cachedValue;

                if(cval != null)
                {
                    List<string>? values = JsonConvert.DeserializeObject<List<string>>(cval);

                    if (values != null)
                    {
                        values.Add(value);

                        return true;
                    }
                }
            } 
            else 
            {
                List<string> values = new List<string>();

                values.Add(value);
                redisDB.StringSet(key, JsonConvert.SerializeObject(values));

                return true;
            }

            return false;
        }

        public bool AddCacheList(string key, List<string> value)
        {
            if (redisDB == null)
            {
                return false;
            }

            RedisValue cachedValue = redisDB.StringGet(key);

            if (!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                string? cval = cachedValue;

                if(cval != null)
                {
                    List<string>? values = JsonConvert.DeserializeObject<List<string>>(cval);

                    if (values != null)
                    {
                        values.AddRange(value);

                        return true;
                    }
                }
            }
            else
            {
                List<string?> values = new List<string?>();

                values.AddRange(value);
                redisDB.StringSet(key, JsonConvert.SerializeObject(values));
                return true;
            }

            return false;
        }

        public bool AddCacheList(string key, List<dynamic> value)
        {
            if (redisDB == null)
            {
                return false;
            }

            RedisValue cachedValue = redisDB.StringGet(key);

            if (!cachedValue.IsNullOrEmpty && cachedValue.HasValue)
            {
                string? cval = cachedValue;
                
                if(cval != null)
                {
                    List<dynamic>? values = JsonConvert.DeserializeObject<List<dynamic>>(cval);

                    if (values != null)
                    {
                        values.AddRange(value);

                        return true;
                    }
                }
            }
            else
            {
                List<dynamic?> values = new List<dynamic?>();

                values.AddRange(value);
                redisDB.StringSet(key, JsonConvert.SerializeObject(values));
                return true;
            }

            return false;
        }
    }
}
