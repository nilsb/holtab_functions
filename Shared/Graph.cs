﻿using Shared.Models;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Logging;

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

        public Graph(Settings _settings)
        {
            if (_settings != null)
            {
                settings = _settings;
                graphClient = _settings.GraphClient;
                log = _settings.log;
                SqlConnectionString = _settings.SqlConnectionString;

                if (!string.IsNullOrEmpty(SqlConnectionString))
                {
                    services = new Services(SqlConnectionString, log);
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

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null)
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


                if (memberToAdd != default(User))
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
                                    {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('" + memberToAdd.Id + "')"}
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

        public async Task<Team?> GetTeamFromGroup(Group group)
        {
            Team? foundTeam = null;

            if (graphClient == null)
            {
                return foundTeam;
            }

            try
            {
                foundTeam = await graphClient.Groups[group.Id].Team.GetAsync();
            }
            catch (Exception)
            {
            }

            return foundTeam;
        }

        public async Task<Team?> GetTeamFromGroup(string groupId)
        {
            Team? foundTeam = null;

            if (graphClient == null)
            {
                return foundTeam;
            }

            try
            {
                foundTeam = await graphClient.Groups[groupId].Team.GetAsync();
            }
            catch (Exception)
            {
            }

            return foundTeam;
        }

        public async Task<Team?> CreateTeamFromGroup(Group group)
        {
            Team? createdTeam = null;

            if(graphClient == null)
            {
                return createdTeam;
            }

            try
            {
                createdTeam = await graphClient.Groups[group.Id].Team.GetAsync();
            }
            catch (Exception)
            {
            }

            if (createdTeam == null)
            {
                log?.LogInformation("Creating team for group " + group.DisplayName);

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
                    createdTeam = await graphClient.Groups[group.Id].Team.PutAsync(teamSettings);

                    log?.LogInformation("Waiting 60s for team to be created");
                    //wait for team to be created
                    Thread.Sleep(60000);
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.Message);
                }
            }

            return createdTeam;
        }

        public async Task<TeamsApp?> AddTeamApp(Team team, string appId)
        {
            TeamsApp? returnValue = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                var result = await graphClient.Teams[team.Id].InstalledApps.GetAsync(request => { request.QueryParameters.Expand = new string[] { "teamsApp" }; });

                if (result?.Value?.Count > 0)
                {

                    foreach (var app in result.Value)
                    {
                        if (app.TeamsApp?.Id == appId)
                        {
                            returnValue = app.TeamsApp;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.Message);
            }

            if (returnValue == null)
            {
                try
                {
                    log?.LogInformation("Add app to team " + team.DisplayName);
                    var teamsAppInstallation = new TeamsAppInstallation
                    {
                        AdditionalData = new Dictionary<string, object>()
                    {
                        {"teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + appId}
                    }
                    };

                    var installation = await graphClient.Teams[team.Id].InstalledApps.PostAsync(teamsAppInstallation);

                    if (installation != null)
                    {
                        returnValue = installation.TeamsApp;
                    }

                }
                catch (Exception ex)
                {
                    log?.LogError(ex.Message);
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

        public async Task<bool> TabExists(Team team, Channel channel, string tabName)
        {
            bool returnValue = false;

            if(graphClient == null)
            {
                return false;
            }

            try
            {
                var tabs = await graphClient.Teams[team.Id].Channels[channel.Id].Tabs.GetAsync();

                if(tabs?.Value?.Count > 0)
                {
                    returnValue = tabs.Value.Any(t => t.DisplayName == tabName);
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.Message);
            }

            return returnValue;
        }

        public async Task<bool> AddChannelApp(Team team, TeamsApp app, Channel channel, string tabName, string entityId, string contentUrl, string webUrl, string removeUrl)
        {
            bool returnValue = false;

            if(!(await TabExists(team, channel, tabName)))
            {
                try
                {
                    log?.LogInformation("Adding app tab");

                    TeamsTab infotab = new TeamsTab()
                    {
                        DisplayName = tabName,
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "teamsApp@odata.bind", "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + app.Id }
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
                        var tab = await graphClient.Teams[team.Id].Channels[channel.Id].Tabs.PostAsync(infotab);
                        returnValue = true;
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.Message);
                }
            }

            return returnValue;
        }

        public async Task<Channel?> FindChannel(Team team, string channelName)
        {
            Channel? returnValue = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                log?.LogInformation("Find channel " + channelName + " in team " + team.DisplayName);

                var channels = await graphClient.Teams[team.Id].Channels.GetAsync();

                if(channels?.Value?.Count() > 0)
                {
                    returnValue = channels.Value.FirstOrDefault(c => c.DisplayName == channelName);
                    log?.LogInformation("Channel " + channelName + " found in team " + team.DisplayName);
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.Message);
            }

            return returnValue;
        }

        public async Task<Channel?> AddChannel(Team team, string channelName, string channelDescription, ChannelMembershipType type)
        {
            Channel? returnValue = null;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                log?.LogInformation("Add channel " + channelName + " to team " + team.DisplayName);
                var channel = new Channel
                {
                    DisplayName = channelName,
                    Description = channelDescription,
                    MembershipType = type
                };

                returnValue = await graphClient.Teams[team.Id].Channels.PostAsync(channel);
            }
            catch (Exception ex)
            {
                log?.LogError(ex.Message);
            }

            return returnValue;
        }

        public async Task CreatePlannerTabInChannelAsync(string teamId, string tabName, string channelId, string planId)
        {
            if(settings == null || settings.GraphClient == null)
            {
                return;
            }

            var tab = new TeamsTab
            {
                DisplayName = tabName,
                TeamsApp = new TeamsApp
                {
                    Id = "com.microsoft.teamspace.tab.planner"
                },
                Configuration = new TeamsTabConfiguration
                {
                    EntityId = planId,
                    ContentUrl = $"https://tasks.office.com/{settings.TenantID}/en-US/Home/PlannerFrame?page=7&planId={planId}&auth=true",
                    WebsiteUrl = $"https://tasks.office.com/{settings.TenantID}/en-US/Home/PlanViews/{planId}",
                    RemoveUrl = $"https://tasks.office.com/{settings.TenantID}/en-US/Home/PlannerFrame?page=13&planId={planId}&auth=true"
                }
            };

            try
            {
                var createdTab = await settings.GraphClient.Teams[teamId].Channels[channelId].Tabs.PostAsync(tab);

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
        public async Task<bool> AddGroupOwner(string userEmail, string GroupId)
        {
            User? memberToAdd = default(User);
            bool returnValue = true;

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null)
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

                if (memberToAdd != default(User))
                {
                    var directoryObject = new ReferenceCreate
                    {
                        OdataId = "https://graph.microsoft.com/v1.0/directoryObjects/" + memberToAdd.Id
                    };

                    bool ownerExists = false;

                    try
                    {
                        var owners = await graphClient.Groups[GroupId].Owners.GetAsync();

                        if (owners?.Value?.Count > 0)
                        {
                            foreach (var owner in owners.Value)
                            {
                                if (owner.Id == memberToAdd.Id)
                                {
                                    ownerExists = true;
                                }
                            }

                            if (!ownerExists)
                            {
                                log?.LogInformation($"Adding owner {userEmail}: {memberToAdd.Id} to group");
                                await graphClient.Groups[GroupId].Owners.Ref.PostAsync(directoryObject);
                                ownerExists = true;
                            }
                        }
                        else
                        {
                            log?.LogInformation($"Adding owner {userEmail}: {memberToAdd.Id} to group");
                            await graphClient.Groups[GroupId].Owners.Ref.PostAsync(directoryObject);
                            ownerExists = true;
                        }

                        returnValue = true;
                    }
                    catch (Exception ex)
                    {
                        log?.LogError(ex.Message);
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

        public async Task<bool> AddGroupMember(string userEmail, string GroupId)
        {
            User? memberToAdd = default(User);
            bool returnValue = true;

            if (!string.IsNullOrEmpty(userEmail) && graphClient != null)
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

                if (memberToAdd != default(User))
                {
                    var directoryObject = new ReferenceCreate
                    {
                        OdataId = "https://graph.microsoft.com/v1.0/directoryObjects/" + memberToAdd.Id
                    };

                    bool memberExists = false;

                    try
                    {
                        var members = await graphClient.Groups[GroupId].Members.GetAsync();

                        if(members?.Value?.Count > 0)
                        {
                            foreach(var member in members.Value)
                            {
                                if(member.Id == memberToAdd.Id)
                                {
                                    memberExists = true;
                                }
                            }

                            if(!memberExists)
                            {
                                log?.LogInformation($"Adding owner {userEmail}: {memberToAdd.Id} to group");
                                await graphClient.Groups[GroupId].Members.Ref.PostAsync(directoryObject);
                                memberExists = true;
                            }
                        }
                        else
                        {
                            log?.LogInformation($"Adding owner {userEmail}: {memberToAdd.Id} to group");
                            await graphClient.Groups[GroupId].Members.Ref.PostAsync(directoryObject);
                            memberExists = true;
                        }

                        returnValue = true;
                    }
                    catch (Exception ex)
                    {
                        log?.LogError(ex.Message);
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

        public async Task<Group?> CreateGroup(string description, string mailNickname, List<string> owners)
        {
            Group? createdGroup = default(Group);

            if (!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname) && graphClient != null)
            {
                log?.LogInformation("Creating group " + description);
                log?.LogInformation("With owners: " + String.Join(",", owners));

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
                    log?.LogInformation("Error creating group.");
                    log?.LogError(ex.Message);
                }

                log?.LogInformation("Waiting 60s for group to be created.");

                //wait for group to be created
                Thread.Sleep(90000);
            }

            return createdGroup;
        }

        public async Task<Group?> CreateGroupNoWait(string description, string mailNickname, List<string> owners)
        {
            Group? createdGroup = default(Group);

            if (!string.IsNullOrEmpty(description) && !string.IsNullOrEmpty(mailNickname) && graphClient != null)
            {
                log?.LogInformation("Creating group " + description);
                log?.LogInformation("With owners: " + String.Join(",", owners));

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
                    log?.LogInformation("Error creating group.");
                    log?.LogError(ex.Message);
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

        public async Task<FindGroupResult> GetGroupById(string id)
        {
            FindGroupResult returnValue = new FindGroupResult();
            returnValue.Success = false;

            if(graphClient == null)
            {
                return returnValue;
            }

            try
            {
                Group? group = await graphClient.Groups[id].GetAsync();

                if (group != null)
                {
                    returnValue.Success = true;
                    returnValue.Count = 1;
                    returnValue.group = group;
                }
            }
            catch (Exception ex)
            {
                log?.LogInformation($"Group not found. {ex}");
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
                returnValue.groups = foundGroups.Value;
            }
            else if(foundGroups?.Value?.Count > 0)
            {
                returnValue.Success = true;
                returnValue.Count = 1;
                returnValue.group = foundGroups.Value[0];
            }
            else
            {
                returnValue.Success = false;
            }

            return returnValue;
        }
        #endregion

        #region Drives
        public async Task<Drive?> GetGroupDrive(Group? group)
        {
            Drive? groupDrive = null;

            if(graphClient == null || group == null)
            {
                return null;
            }

            try
            {
                groupDrive = await graphClient.Groups[group.Id].Drive.GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }

            return groupDrive;
        }

        public async Task<Drive?> GetGroupDrive(string? GroupId)
        {
            Drive? groupDrive = null;

            if (graphClient == null || string.IsNullOrEmpty(GroupId))
            {
                return null;
            }

            try
            {
                groupDrive = await graphClient.Groups[GroupId].Drive.GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }

            return groupDrive;
        }

        public async Task<Site?> GetGroupSite(string? GroupId)
        {
            Site? returnValue = null;

            if (graphClient == null || string.IsNullOrEmpty(GroupId))
            {
                return null;
            }

            FindGroupResult? findGroup = await GetGroupById(GroupId);

            if(findGroup?.Success == true && findGroup?.group != null)
            {
                var sites = await graphClient.Groups[findGroup.group.Id].Sites.GetAsync();

                if(sites?.Value?.Count > 0)
                {
                    returnValue = sites?.Value[0];
                }
            }

            return returnValue;
        }

        public async Task<Drive?> GetSiteDrive(Site? site)
        {
            Drive? groupDrive = null;

            if (graphClient == null || site == null)
            {
                return null;
            }

            try
            {
                groupDrive = await graphClient.Sites[site.Id].Drive.GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }

            return groupDrive;
        }

        public async Task<Drive?> GetSiteDrive(string? SiteId)
        {
            Drive? groupDrive = null;

            if (graphClient == null || string.IsNullOrEmpty(SiteId))
            {
                return null;
            }

            try
            {
                groupDrive = await graphClient.Sites[SiteId].Drive.GetAsync();
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }

            return groupDrive;
        }

        public async Task<List<DriveItem>> GetDriveRootItems(Drive? groupDrive)
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
                    log?.LogError(ex.ToString());
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
                        log?.LogError(ex.ToString());
                    }

                    if(rootChildren?.Value?.Count > 0)
                    {
                        returnValue = rootChildren.Value.ToList();
                    }
                }
            }

            return returnValue;
        }

        public async Task<List<DriveItem>> GetDriveFolderChildren(Drive? groupDrive, DriveItem? parent, bool recursive = false)
        {
            List<DriveItem> returnValue = new List<DriveItem>();

            if (groupDrive != null && graphClient != null && parent != null)
            {
                var folderChildren = await graphClient.Drives[groupDrive.Id].Items[parent.Id].Children.GetAsync();

                if (folderChildren?.Value?.Count > 0)
                {
                    if (recursive)
                    {
                        foreach(var child in folderChildren.Value)
                        {
                            var subchildren = await GetDriveFolderChildren(groupDrive, child, recursive);

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

        public async Task<DriveItem?> FindItem(Drive? groupDrive, string? Path, bool withRetry = false)
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
            catch (Exception)
            {
            }

            while (returnValue == null && withRetry)
            {
                if(returnValue != null)
                {
                    lock (returnValue)
                    {
                        log?.LogInformation($"Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");
                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDrive.Id].Root.ItemWithPath(Path).GetAsync();
                }
                catch (Exception)
                {
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DriveItem?> FindItem(Drive? groupDrive, string? parentId, string? Path, bool withRetry)
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
            catch (Exception)
            {
            }

            int maxcnt = 2;
            int cnt = 0;

            while (returnValue == null && withRetry)
            {
                if(returnValue != null)
                {
                    lock (returnValue)
                    {
                        log?.LogInformation($"Item not found, trying again... (attempt {cnt + 1} of {maxcnt + 1})");
                        cnt += 1;
                        Thread.Sleep(10000);
                    }
                }

                try
                {
                    returnValue = await graphClient.Drives[groupDrive.Id].Items[parentId].ItemWithPath(Path).GetAsync();
                }
                catch (Exception)
                {
                }

                if (cnt == maxcnt)
                    break;
            }

            return returnValue ?? default(DriveItem);
        }

        public async Task<DownloadFileResult> DownloadFile(string? GroupID, string? FolderID, string? FileName)
        {
            DownloadFileResult returnValue = new DownloadFileResult();
            Stream orderFileStream = Stream.Null;

            if(graphClient == null || string.IsNullOrEmpty(GroupID) || string.IsNullOrEmpty(FolderID) || string.IsNullOrEmpty(FileName))
            {
                return returnValue;
            }

            try
            {
                Drive? groupDrive = await GetGroupDrive(GroupID);

                if(groupDrive != null)
                {
                    //download order file content
                    returnValue.Contents = await graphClient.Drives[groupDrive.Id].Items[FolderID].ItemWithPath(FileName).Content.GetAsync() ?? Stream.Null;
                    returnValue.Success = true;
                }
            }
            catch (Exception)
            {
                returnValue.Success = false;
            }

            return returnValue;
        }

        public async Task<DownloadFileResult> DownloadFile(Group? Group, DriveItem? Folder, string? Path)
        {
            DownloadFileResult returnValue = new DownloadFileResult();
            Stream orderFileStream = Stream.Null;

            if (graphClient == null || Group == null || Folder == null || string.IsNullOrEmpty(Path))
            {
                return returnValue;
            }


            try
            {
                Drive? groupDrive = await GetGroupDrive(Group);

                if(groupDrive != null)
                {
                    //download order file content
                    returnValue.Contents = await graphClient.Drives[groupDrive.Id].Items[Folder.Id].ItemWithPath(Path).Content.GetAsync() ?? Stream.Null;
                    returnValue.Success = true;
                }
            }
            catch (Exception)
            {
                returnValue.Success = false;
            }

            return returnValue;
        }

        public async Task<bool> UploadFile(string? GroupID, string? FolderID, string? FileName, Stream? FileContents)
        {
            bool returnValue = false;
            Stream fileStream = Stream.Null;

            if (graphClient == null || FileContents == null || FileContents == Stream.Null || string.IsNullOrEmpty(GroupID) || string.IsNullOrEmpty(FolderID) || string.IsNullOrEmpty(FileName))
            {
                return returnValue;
            }

            try
            {
                Drive? groupDrive = await GetGroupDrive(GroupID);

                if(groupDrive != null)
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
                    var fileUploadSession = await graphClient.Drives[groupDrive.Id].Items[FolderID].ItemWithPath(FileName).CreateUploadSession.PostAsync(uploadRequest);

                    if (fileUploadSession != null)
                    {
                        var fileUploadTask = new LargeFileUploadTask<DriveItem>(fileUploadSession, fileStream, maxUploadChunkSize, graphClient.RequestAdapter);

                        var totalLength = fileStream.Length;
                        // Create a callback that is invoked after each slice is uploaded
                        IProgress<long> progress = new Progress<long>(prog => {
                            log?.LogInformation($"Uploaded {prog} bytes of {totalLength} bytes");
                        });

                        // Upload the file
                        var uploadResult = await fileUploadTask.UploadAsync(progress);

                        log?.LogInformation(uploadResult.UploadSucceeded ?
                            $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                            "Upload failed");
                        returnValue = true;
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError($"Error uploading: {ex.ToString()}");
                returnValue = false;
            }

            return returnValue;
        }

        public async Task<bool> UploadFile(Group? Group, DriveItem? Folder, string? Path, Stream? FileContents)
        {
            bool returnValue = false;
            Stream fileStream = Stream.Null;

            if(graphClient == null || Group == null || Folder == null || string.IsNullOrEmpty(Path) || FileContents == null || FileContents == Stream.Null)
            {
                return returnValue;
            }

            try
            {
                Drive? groupDrive = await GetGroupDrive(Group.Id);

                if (groupDrive != null)
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
                    var fileUploadSession = await graphClient.Drives[groupDrive.Id].Items[Folder.Id].ItemWithPath(Path).CreateUploadSession.PostAsync(uploadRequest);

                    if (fileUploadSession != null)
                    {
                        var fileUploadTask = new LargeFileUploadTask<DriveItem>(fileUploadSession, fileStream, maxUploadChunkSize, graphClient.RequestAdapter);

                        var totalLength = fileStream.Length;
                        // Create a callback that is invoked after each slice is uploaded
                        IProgress<long> progress = new Progress<long>(prog => {
                            log?.LogInformation($"Uploaded {prog} bytes of {totalLength} bytes");
                        });

                        // Upload the file
                        var uploadResult = await fileUploadTask.UploadAsync(progress);

                        log?.LogInformation(uploadResult.UploadSucceeded ?
                            $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                            "Upload failed");
                        returnValue = true;
                    }
                }
            }
            catch (Exception)
            {
                returnValue = false;
            }

            return returnValue;
        }

        public async Task<bool> CopyFile(CopyItem? source, CopyItem? destination)
        {
            bool returnValue = false;

            if(source != null && destination != null && !string.IsNullOrEmpty(source.GroupId) && !string.IsNullOrEmpty(source.FolderId) && !string.IsNullOrEmpty(source.Path))
            {
                //download the file
                log?.LogInformation("Download file " + source.Path);
                DownloadFileResult downloadFile = await this.DownloadFile(source.GroupId, source.FolderId, source.Path);

                if (downloadFile.Success && !string.IsNullOrEmpty(destination.GroupId) && !string.IsNullOrEmpty(destination.FolderId) && !string.IsNullOrEmpty(destination.Path))
                {
                    log?.LogInformation("Upload file " + destination.Path);
                    if (await this.UploadFile(destination.GroupId, destination.FolderId, destination.Path, downloadFile.Contents))
                    {
                        returnValue = true;
                    }
                }
            }

            return returnValue;
        }

        public async Task<bool> MoveFile(CopyItem? source, CopyItem? destination)
        {
            bool returnValue = false;

            if (graphClient != null && source != null && destination != null && !string.IsNullOrEmpty(source.GroupId) && !string.IsNullOrEmpty(source.FolderId) && !string.IsNullOrEmpty(source.Path))
            {
                //download the file
                DownloadFileResult downloadFile = await this.DownloadFile(source.GroupId, source.FolderId, source.Path);

                if (downloadFile.Success && !string.IsNullOrEmpty(destination.GroupId) && !string.IsNullOrEmpty(destination.FolderId) && !string.IsNullOrEmpty(destination.Path))
                {
                    if (await this.UploadFile(destination.GroupId, destination.FolderId, destination.Path, downloadFile.Contents))
                    {
                        try
                        {
                            Drive? groupDrive = await GetGroupDrive(source.GroupId);

                            if(groupDrive != null)
                            {
                                await graphClient.Drives[groupDrive.Id].Items[source.FileId].DeleteAsync();
                                returnValue = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CopyFolder(string? GroupId, string? ParentId, DriveItem? Folder, bool recursive = false, bool includeFiles = false)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;

            if(string.IsNullOrEmpty(GroupId) || string.IsNullOrEmpty(ParentId) || Folder == null || string.IsNullOrEmpty(Folder.Name))
            {
                return returnValue;
            }

            DriveItem? createdFolder = null;
            CreateFolderResult result = await this.CreateFolder(GroupId, ParentId, Folder.Name);

            if (result.Success && Folder.Children != null)
            {
                log?.LogInformation("Created " + Folder.Name + " folder.");

                createdFolder = result.folder;

                if (recursive && createdFolder != null)
                {
                    createdFolder.Children = new List<DriveItem>();

                    foreach (var childFolder in Folder.Children)
                    {
                        if (childFolder.Folder == null)
                            continue;

                        var createdChild = await this.CopyFolder(GroupId, createdFolder.Id, childFolder, recursive, includeFiles);

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
                        await this.CopyFile(source, destination);
                    }
                }

                returnValue.folder = createdFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CopyFolder(string GroupId, DriveItem Folder, bool recursive, bool? includeFiles)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;
            log?.LogInformation("Creating " + Folder.Name + " folder.");
            DriveItem? createdFolder = null;
            CreateFolderResult result = new CreateFolderResult();
            result.Success = false;

            if (!string.IsNullOrEmpty(Folder.Name))
            {
                result = this.CreateFolder(GroupId, Folder.Name).Result;
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

                        var createdChild = await this.CopyFolder(GroupId, childFolder, recursive, includeFiles);

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
                        await this.CopyFile(source, destination);
                    }
                }

                returnValue.folder = createdFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CreateFolder(string GroupId, string ParentId, string FolderName)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;
            DriveItem? createdFolder = null;

            //first check if folder exists
            var drive = this.GetGroupDrive(GroupId).Result;
            var existingFolder = this.FindItem(drive, ParentId, FolderName, false).Result;

            if(existingFolder == null && graphClient != null && drive != null)
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

                    createdFolder = await graphClient.Drives[drive.Id].Items[ParentId].Children.PostAsync(driveItemFolder);

                    if (createdFolder != null)
                    {
                        returnValue.folder = createdFolder;
                        returnValue.Success = true;
                        returnValue.Existed = false;
                        log?.LogInformation("Created " + FolderName + " folder.");
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.ToString());
                }
            }
            else
            {
                returnValue.Existed = true;
                returnValue.folder = existingFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }

        public async Task<CreateFolderResult> CreateFolder(string GroupId, string FolderName)
        {
            CreateFolderResult returnValue = new CreateFolderResult();
            returnValue.Success = false;
            DriveItem? createdFolder = null;

            //first check if folder exists
            var drive = this.GetGroupDrive(GroupId).Result;
            var existingFolder = this.FindItem(drive, FolderName, false).Result;

            if (existingFolder == null && drive != null && graphClient != null)
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

                    var rootItem = await graphClient.Drives[drive.Id].Root.GetAsync();
                    
                    if(rootItem != null)
                    {
                        createdFolder = await graphClient.Drives[drive.Id].Items[rootItem.Id].Children.PostAsync(driveItemFolder);

                        if (createdFolder != null)
                        {
                            returnValue.folder = createdFolder;
                            returnValue.Success = true;
                            log?.LogInformation("Created " + FolderName + " folder.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.ToString());
                }
            }
            else
            {
                returnValue.folder = existingFolder;
                returnValue.Success = true;
            }

            return returnValue;
        }
        #endregion

        #region List
        public async Task<List<ListItem>> GetListItems(string SiteId, string ListId, string Filter)
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

        public async Task CreateDriveColumn(string groupId, ColumnDefinition def)
        {
            if(graphClient == null)
            {
                return;
            }

            try
            {
                var group = await GetGroupById(groupId);

                if(group.Success)
                {
                    var drive = await GetGroupDrive(groupId);

                    if(drive != null)
                    {
                        var list = await graphClient.Drives[drive.Id].List.GetAsync();
                        var site = await GetGroupSite(groupId);

                        if(list != null && site != null)
                        {
                            log?.LogInformation($"Adding column: {def.Description}");
                            var col = await graphClient.Sites[site.Id].Lists[list.Id].Columns.PostAsync(def);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }
        }

        public async Task CreateDriveColumn(Site site, List list, ColumnDefinition def)
        {
            if(graphClient == null)
            {
                return;
            }

            try
            {
                log?.LogInformation($"Adding column: {def.Description}");
                var col = await graphClient.Sites[site.Id].Lists[list.Id].Columns.PostAsync(def);
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
            }
        }
        #endregion

        #region Plans
        public async Task<PlannerPlan?> CreatePlanAsync(string groupId, string planName)
        {
            PlannerPlan? createdPlan = null;

            if (settings == null || settings.GraphClient == null)
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
                createdPlan = await settings.GraphClient.Planner.Plans
                    .PostAsync(newPlan);

                log?.LogInformation($"Plan created with ID: {createdPlan?.Id}");
            }
            catch (ServiceException ex)
            {
                log?.LogError($"Error creating plan: {ex.Message}");
            }

            return createdPlan;
        }

        public async Task<IList<PlannerPlan>> GetPlansAsync(string groupId)
        {
            List<PlannerPlan> returnValue = new List<PlannerPlan>();

            if(settings == null || settings.GraphClient == null)
            {
                return returnValue;
            }

            try
            {
                var plans = await settings.GraphClient.Groups[groupId].Planner.Plans.GetAsync();

                if (plans?.Value?.Count > 0)
                {
                    services?.Log("Found: " + plans.Value.Count + " plans in group");

                    foreach (PlannerPlan plan in plans.Value)
                    {
                        returnValue.Add(plan);
                    }
                }
            }
            catch (ServiceException ex)
            {
                log?.LogError($"Error retrieving plans: {ex.Message}");
            }

            return returnValue;
        }

        public async Task<PlannerPlan?> PlanExists(string groupId, string planTitle)
        {
            services?.Log("Trying to find plan: " + planTitle + " in group: " + groupId);
            var plans = await GetPlansAsync(groupId);

            if(plans?.Count > 0)
            {
                services?.Log("Found: " + plans.Count);
                if (plans.Any(p => p.Title == planTitle))
                {
                    return plans.FirstOrDefault(p => p.Title == planTitle);
                }
            }

            return null;
        }

        public async Task<IList<PlannerBucket>> GetBucketsAsync(string planId)
        {
            List<PlannerBucket> returnValue = new List<PlannerBucket>();

            if (settings == null || settings.GraphClient == null)
            {
                return returnValue;
            }

            try
            {
                var buckets = await settings.GraphClient.Planner.Plans[planId].Buckets
                    .GetAsync();

                if(buckets?.Value?.Count > 0)
                {
                    returnValue = buckets.Value;
                }
            }
            catch (ServiceException ex)
            {
                log?.LogError($"Error retrieving buckets: {ex.Message}");
            }

            return returnValue;
        }

        public async Task CopyBucketAsync(PlannerBucket sourceBucket, string targetPlanId)
        {
            if(settings == null || settings.GraphClient == null)
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
                createdBucket = await settings.GraphClient.Planner.Buckets.PostAsync(newBucket);
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating bucket: {ex.Message}");
                return;
            }

            if(createdBucket != null)
            {
                // Retrieve tasks from the source bucket
                var tasks = await settings.GraphClient.Planner.Buckets[sourceBucket.Id].Tasks
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
                            await settings.GraphClient.Planner.Tasks
                                .PostAsync(newTask);
                        }
                        catch (ServiceException ex)
                        {
                            log?.LogError($"Error creating task: {ex.Message}");
                        }
                    }
                }
            }
        }

        #endregion
    }
}
