/**
 * API response fixtures for testing.
 * 
 * These are based on real API response structures documented in docs/API-REFERENCE.md.
 * They represent the shape of data returned by Teams/Substrate APIs.
 */

/**
 * Substrate v2 search result item.
 * From: POST https://substrate.office.com/searchservice/api/v2/query
 */
export const searchResultItem = {
  Id: 'AAMkAGE1OWFlZjc0LWYxMjQtNGM1Mi05NzJlLTU0MTU2ZGU1OGM1YQBGAAAAAACaT2h4EH4ZT5pQgKA-example',
  ReferenceId: 'abc123-def456.1000.1',
  HitHighlightedSummary: 'Let me check the <c0>budget</c0> report for Q3',
  Summary: 'Let me check the budget report for Q3',
  Source: {
    DateTimeReceived: '2026-01-20T14:30:00.000Z', // = 1768919400000
    From: {
      EmailAddress: {
        Name: 'Smith, John',
        Address: 'john.smith@company.com',
      },
    },
    ChannelName: 'General',
    TeamName: 'Finance Team',
    Extensions: {
      SkypeSpaces_ConversationPost_Extension_SkypeGroupId: '19:abcdef123456@thread.tacv2',
    },
    // Top-level post: messageid matches DateTimeReceived timestamp
    ClientConversationId: '19:abcdef123456@thread.tacv2;messageid=1768919400000',
  },
};

/**
 * Search result with HTML content that needs stripping.
 */
export const searchResultWithHtml = {
  Id: 'AAMkBGFiY2RlZg',
  ReferenceId: 'xyz789.1000.1',
  HitHighlightedSummary: '<p>Meeting <strong>notes</strong> from &amp; yesterday&apos;s call</p><br/><div>Action items:</div>',
  Source: {
    DateTimeReceived: '2026-01-21T09:00:00.000Z',
    From: 'Jane Doe',
    ClientThreadId: '19:meeting123@thread.v2',
  },
};

/**
 * Minimal search result with only required fields.
 */
export const searchResultMinimal = {
  Id: 'minimal-id',
  HitHighlightedSummary: 'A short message here',
};

/**
 * Search result too short to be valid (content < 5 chars).
 */
export const searchResultTooShort = {
  Id: 'short-id',
  HitHighlightedSummary: 'Hi',
};

/**
 * Full EntitySets response structure from v2 query.
 */
export const searchEntitySetsResponse = {
  EntitySets: [
    {
      ResultSets: [
        {
          Total: 4307,
          Results: [
            searchResultItem,
            searchResultWithHtml,
          ],
        },
      ],
    },
  ],
};

/**
 * Person suggestion from Substrate suggestions API.
 * From: POST https://substrate.office.com/search/api/v1/suggestions
 */
export const personSuggestion = {
  Id: 'a1b2c3d4-e5f6-7890-abcd-ef1234567890@company.onmicrosoft.com',
  MRI: '8:orgid:a1b2c3d4-e5f6-7890-abcd-ef1234567890',
  DisplayName: 'Smith, John',
  GivenName: 'John',
  Surname: 'Smith',
  EmailAddresses: ['john.smith@company.com'],
  CompanyName: 'Acme Corp',
  Department: 'Engineering',
  JobTitle: 'Senior Engineer',
};

/**
 * Person with minimal info using a proper GUID (no optional fields).
 */
export const personMinimal = {
  Id: 'b1c2d3e4-f5a6-7890-bcde-1234567890ab',
  DisplayName: 'Jane Doe',
};

/**
 * Person with base64-encoded GUID (real-world API response format).
 * The base64 '93qkaTtFGWpUHjyRafgdhg==' decodes to GUID '69a47af7-453b-6a19-541e-3c9169f81d86'.
 */
export const personWithBase64Id = {
  Id: '93qkaTtFGWpUHjyRafgdhg==',
  MRI: '8:orgid:93qkaTtFGWpUHjyRafgdhg==',
  DisplayName: 'Rob MacDonald',
  EmailAddresses: ['rob@company.com'],
};

/**
 * Groups response from suggestions API.
 */
export const peopleGroupsResponse = {
  Groups: [
    {
      Suggestions: [
        personSuggestion,
        personMinimal,
      ],
    },
  ],
};

/**
 * JWT payload with full user info.
 */
export const jwtPayloadFull = {
  oid: 'user-object-id-guid',
  name: 'Macdonald, Rob',
  upn: 'rob.macdonald@company.com',
  preferred_username: 'rob@company.com',
  email: 'rob.m@personal.com',
  given_name: 'Rob',
  family_name: 'Macdonald',
  tid: 'tenant-id-guid',
  exp: 1705850000,
  iat: 1705846400,
};

/**
 * JWT payload with minimal info (only required fields).
 */
export const jwtPayloadMinimal = {
  oid: 'another-user-guid',
  name: 'Alice Smith',
  exp: 1705850000,
};

/**
 * JWT payload for name parsing tests - "Surname, GivenName" format.
 */
export const jwtPayloadCommaName = {
  oid: 'comma-name-user',
  name: 'Jones, David',
  upn: 'david.jones@company.com',
};

/**
 * JWT payload for name parsing tests - "GivenName Surname" format.
 */
export const jwtPayloadSpaceName = {
  oid: 'space-name-user',
  name: 'Sarah Connor',
  upn: 'sarah.connor@company.com',
};

/**
 * Channel thread reply - messageid differs from message timestamp.
 * The parent post was at 1768919400000 (2026-01-20T14:30:00.000Z),
 * this reply is at 1768921200000 (2026-01-20T15:00:00.000Z).
 */
export const searchResultThreadReply = {
  Id: 'AAMkThreadReply',
  ReferenceId: 'thread-reply.1000.1',
  HitHighlightedSummary: 'Thanks for the update on the budget!',
  Source: {
    DateTimeReceived: '2026-01-20T15:00:00.000Z', // = 1768921200000
    From: {
      EmailAddress: {
        Name: 'Doe, Jane',
        Address: 'jane.doe@company.com',
      },
    },
    ChannelName: 'General',
    TeamName: 'Finance Team',
    Extensions: {
      SkypeSpaces_ConversationPost_Extension_SkypeGroupId: '19:abcdef123456@thread.tacv2',
    },
    // Parent message ID (1768919400000) differs from this message's timestamp (1768921200000)
    ClientConversationId: '19:abcdef123456@thread.tacv2;messageid=1768919400000',
    ClientThreadId: '19:abcdef123456@thread.tacv2',
  },
};

/**
 * Message source with explicit message ID.
 */
export const sourceWithMessageId = {
  MessageId: '1705760000000',
  DateTimeReceived: '2026-01-20T12:00:00.000Z',
  ClientConversationId: '19:thread@tacv2',
};

/**
 * Message source with ID in ClientConversationId.
 */
export const sourceWithConvIdMessageId = {
  DateTimeReceived: '2026-01-20T12:00:00.000Z',
  ClientConversationId: '19:thread@tacv2;messageid=1705770000000',
};

/**
 * Email search result from Substrate v2 query API.
 * From: POST https://substrate.office.com/searchservice/api/v2/query
 * with entityType: 'Message', contentSources: ['Exchange']
 */
export const emailSearchResult = {
  Id: 'AAMkAGE1OWFlZjc0LWYxMjQtNGM1Mi05NzJlLTU0MTU2ZGU1OGM1YQBGAAAAAAEmail',
  Subject: 'Q3 Budget Review',
  HitHighlightedSummary: 'Please review the attached <c0>budget</c0> spreadsheet for Q3',
  Source: {
    Subject: 'Q3 Budget Review',
    From: {
      EmailAddress: {
        Name: 'Smith, John',
        Address: 'john.smith@company.com',
      },
    },
    DateTimeReceived: '2026-01-20T14:30:00.000Z',
    HasAttachments: true,
    Importance: 'Normal',
    IsRead: true,
    DisplayTo: 'Jane Doe; Rob MacDonald',
    DisplayCc: 'Finance Team',
    WebLink: 'https://outlook.office.com/mail/id/AAMkAGE1OWFlZjc0',
    ConversationId: { Id: 'AAQkAGE1OWFlZjc0LWYxMjQtNGM1Mi05NzJl' },
    Preview: 'Please review the attached budget spreadsheet for Q3',
  },
};

/**
 * Email search result with minimal fields.
 */
export const emailSearchResultMinimal = {
  Id: 'AAMkMinimalEmail',
  HitHighlightedSummary: 'Quick update on the project status',
  Source: {
    Subject: 'Project Update',
    FromName: 'Jane Doe',
    FromAddress: 'jane.doe@company.com',
    DateTimeSent: '2026-01-21T09:00:00.000Z',
  },
};

/**
 * Email search result with structured ToRecipients array.
 */
export const emailSearchResultWithRecipients = {
  Id: 'AAMkRecipientsEmail',
  Subject: 'Team Meeting Notes',
  HitHighlightedSummary: 'Here are the notes from today\'s meeting',
  Source: {
    Subject: 'Team Meeting Notes',
    From: 'Alice Johnson',
    DateTimeReceived: '2026-01-22T16:00:00.000Z',
    ToRecipients: [
      { EmailAddress: { Name: 'Bob Wilson', Address: 'bob@company.com' } },
      { EmailAddress: { Name: 'Carol Davis', Address: 'carol@company.com' } },
    ],
    CcRecipients: [
      { EmailAddress: { Name: 'Dave Brown', Address: 'dave@company.com' } },
    ],
    HasAttachments: false,
    IsRead: false,
    Importance: 'High',
  },
};

/**
 * Email search result that is a calendar response (should be filtered by default).
 */
export const emailSearchResultCalendarResponse = {
  Id: 'AAMkCalendarResponse',
  Subject: 'Accepted: Weekly Standup',
  HitHighlightedSummary: '',
  Source: {
    Subject: 'Accepted: Weekly Standup',
    From: {
      EmailAddress: {
        Name: 'Macdonald, Rob',
        Address: 'rob.macdonald@company.com',
      },
    },
    DateTimeReceived: '2026-02-06T12:50:45.000Z',
    HasAttachments: false,
    Importance: 'Normal',
    IsRead: true,
  },
};

/**
 * Email EntitySets response structure from v2 query.
 */
export const emailEntitySetsResponse = {
  EntitySets: [
    {
      ResultSets: [
        {
          Total: 142,
          Results: [
            emailSearchResult,
            emailSearchResultMinimal,
            emailSearchResultWithRecipients,
            emailSearchResultCalendarResponse,
          ],
        },
      ],
    },
  ],
};

/**
 * Thread message from chatsvc API.
 * From: GET /api/chatsvc/{region}/v1/users/ME/conversations/{id}/messages
 */
export const threadMessage = {
  id: '1705760000000',
  content: '<p>Hello team!</p>',
  messagetype: 'RichText/Html',
  contenttype: 'text',
  from: '8:orgid:user-guid-123',
  imdisplayname: 'John Smith',
  originalarrivaltime: '2026-01-20T12:00:00.000Z',
  composetime: '2026-01-20T11:59:58.000Z',
  clientmessageid: 'client-msg-123',
};

/**
 * Favorites folder response.
 * From: POST /api/csa/{region}/api/v1/teams/users/me/conversationFolders
 */
export const favoritesFolderResponse = {
  folderHierarchyVersion: 1705850000000,
  conversationFolders: [
    {
      id: 'tenant-guid~user-guid~Favorites',
      sortType: 'UserDefinedCustomOrder',
      name: 'Favorites',
      folderType: 'Favorites',
      conversationFolderItems: [
        {
          conversationId: '19:abc@thread.tacv2',
          createdTime: 1705700000000,
          lastUpdatedTime: 1705800000000,
        },
        {
          conversationId: '19:xyz@thread.v2',
          createdTime: 1705600000000,
          lastUpdatedTime: 1705750000000,
        },
      ],
    },
  ],
};

// ============================================================================
// Session state fixtures (all values are FAKE)
// ============================================================================

/** Far-future MSAL expiresOn (seconds since epoch, as a string): year 2100. */
export const FAKE_EXPIRES_FUTURE = '4102444800';
/** Past MSAL expiresOn (seconds since epoch, as a string): year 2001. */
export const FAKE_EXPIRES_PAST = '1000000000';

const FAKE_CONSUMERS_TENANT = '9188040d-6c67-4c5b-b112-36a304b66dad';
const FAKE_PERSONAL_HOME_ACCOUNT = `0000000000000000.${FAKE_CONSUMERS_TENANT}`;

/** Builds a fake MSAL AccessToken localStorage item for the personal fixture. */
function fakePersonalToken(name: string, target: string, secret: string, expiresOn: string) {
  return {
    name,
    value: JSON.stringify({
      credentialType: 'AccessToken',
      homeAccountId: FAKE_PERSONAL_HOME_ACCOUNT,
      realm: FAKE_CONSUMERS_TENANT,
      target,
      secret,
      expiresOn,
    }),
  };
}

/**
 * Teams free (personal account) session: teams.live.com origin with a consumers
 * tenant MSAL cache, opaque access tokens, and the DISCOVER-REGION-GTM config.
 */
export const personalSessionState = {
  cookies: [
    { name: 'skypetoken_asm', value: 'fake-personal-skypetoken', domain: '.teams.live.com', path: '/' },
    { name: 'authtoken', value: 'Bearer%3Dfake-personal-authtoken', domain: '.teams.live.com', path: '/' },
    { name: 'skypetoken_asm', value: 'fake-work-skypetoken', domain: '.teams.microsoft.com', path: '/' },
  ],
  origins: [
    {
      origin: 'https://teams.live.com',
      localStorage: [
        {
          name: 'DISCOVER-REGION-GTM',
          value: JSON.stringify({
            item: {
              chatServiceAfd: 'https://teams.live.com/api/chatsvc/consumer',
              middleTier: 'https://teams.live.com/api/mt',
            },
          }),
        },
        fakePersonalToken(
          'msal-mt-token-old',
          'https://mtsvc.fl.teams.microsoft.com/teams.mt.readwrite',
          'fake-opaque-mt-token-old',
          '4102444700',
        ),
        fakePersonalToken(
          'msal-mt-token-new',
          'https://mtsvc.fl.teams.microsoft.com/teams.mt.readwrite',
          'fake-opaque-mt-token',
          FAKE_EXPIRES_FUTURE,
        ),
        fakePersonalToken(
          'msal-auth-token',
          'https://auth.fl.teams.microsoft.com/teams.auth.readwrite',
          'fake-opaque-auth-token',
          FAKE_EXPIRES_FUTURE,
        ),
        fakePersonalToken(
          'msal-auth-token-expired',
          'https://auth.fl.teams.microsoft.com/teams.auth.readwrite',
          'fake-opaque-auth-token-expired',
          FAKE_EXPIRES_PAST,
        ),
        {
          name: 'msal-refresh-token',
          value: JSON.stringify({
            credentialType: 'RefreshToken',
            homeAccountId: FAKE_PERSONAL_HOME_ACCOUNT,
            clientId: '4b3e8f46-56d3-427f-b1e2-d239b2ea6bca',
            secret: 'fake-refresh-token',
          }),
        },
      ],
    },
  ],
};

/** Work/school session: partitioned DISCOVER-REGION-GTM on teams.microsoft.com. */
export const workSessionState = {
  cookies: [
    { name: 'skypetoken_asm', value: 'fake-work-skypetoken', domain: '.teams.microsoft.com', path: '/' },
    { name: 'authtoken', value: 'Bearer%3Dfake-work-authtoken', domain: '.teams.microsoft.com', path: '/' },
  ],
  origins: [
    {
      origin: 'https://teams.microsoft.com',
      localStorage: [
        {
          name: 'DISCOVER-REGION-GTM',
          value: JSON.stringify({
            item: {
              chatServiceAfd: 'https://teams.microsoft.com/api/chatsvc/amer',
              middleTier: 'https://teams.microsoft.com/api/mt/part/amer-02',
              chatSvcAggAfd: 'https://teams.microsoft.com/api/csa/amer',
            },
          }),
        },
        {
          name: 'msal-work-account',
          value: JSON.stringify({
            credentialType: 'AccessToken',
            homeAccountId: 'fake-oid.11111111-2222-3333-4444-555555555555',
            realm: '11111111-2222-3333-4444-555555555555',
            target: 'https://api.spaces.skype.com/Authorization.ReadWrite',
            secret: 'fake-not-a-jwt',
            expiresOn: FAKE_EXPIRES_FUTURE,
          }),
        },
      ],
    },
  ],
};

/**
 * Personal (Teams free) `POST /api/mt/beta/users/searchUsers` response, keyed by
 * the search term (note the key differs in case from a typed "julia"). Fake values only.
 */
export const consumerSearchUsersResponse = {
  Julia: {
    status: 'Ok',
    userProfiles: [
      {
        displayName: 'Julia Fake',
        mri: '8:live:.cid.fake-julia',
        userPrincipalName: 'julia.fake@example.invalid',
        givenName: 'Julia',
        surname: 'Fake',
        objectId: '00000000-0000-0000-0000-00000000aaaa',
      },
      {
        displayName: 'Julia Second',
        mri: '8:live:.cid.fake-julia-two',
        objectId: '00000000-0000-0000-0000-00000000bbbb',
      },
      { displayName: 'No Mri Person' },
    ],
  },
};
