package graphtutorial;

import com.azure.core.credential.AccessToken;
import com.azure.core.credential.TokenRequestContext;
import com.azure.identity.DeviceCodeCredential;
import com.azure.identity.DeviceCodeCredentialBuilder;
import com.azure.identity.DeviceCodeInfo;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.models.User;
import com.microsoft.graph.options.HeaderOption;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.MessageCollectionPage;
import com.microsoft.graph.requests.MessageCollectionRequest;
import com.microsoft.graph.requests.MessageStreamRequest;
import okhttp3.Request;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Properties;
import java.util.function.Consumer;

public class Graph {
    private static Properties _properties;
    private static DeviceCodeCredential _deviceCodeCredential;
    private static GraphServiceClient<Request> _userClient;

    public static void initializeGraphForUserAuth(Properties properties, Consumer<DeviceCodeInfo> challenge) throws Exception {
        // Ensure properties isn't null
        if (properties == null) {
            throw new Exception("Properties cannot be null");
        }

        _properties = properties;

        final String clientId = properties.getProperty("app.clientId");


        final String tenantId = properties.getProperty("app.tenantId");
        final List<String> graphUserScopes = Arrays
                .asList(properties.getProperty("app.graphUserScopes").split(","));

        _deviceCodeCredential = new DeviceCodeCredentialBuilder()
                .clientId(clientId)
                .tenantId(tenantId)
                .challengeConsumer(challenge)
                .build();

        final TokenCredentialAuthProvider authProvider =
                new TokenCredentialAuthProvider(graphUserScopes, _deviceCodeCredential);


        _userClient = GraphServiceClient.builder()
                .authenticationProvider(authProvider)
                .buildClient();
    }


    public static String getUserToken() throws Exception {
        // Ensure credential isn't null
        if (_deviceCodeCredential == null) {
            throw new Exception("Graph has not been initialized for user auth");
        }



        final String[] graphUserScopes = _properties.getProperty("app.graphUserScopes").split(",");



        final TokenRequestContext context = new TokenRequestContext();
        context.addScopes(graphUserScopes);



        final AccessToken token = _deviceCodeCredential.getToken(context).block();



        return token.getToken();
    }

    public static User getUser() throws Exception {
        // Ensure client isn't null
        if (_userClient == null) {
            throw new Exception("Graph has not been initialized for user auth");
        }

        return _userClient.me()
                .buildRequest()
                .select("displayName,mail,userPrincipalName")
                .get();
    }

    public static MessageCollectionPage getInbox() throws Exception {
        // Ensure client isn't null
        if (_userClient == null) {
            throw new Exception("Graph has not been initialized for user auth");
        }

        return _userClient.me()
                .mailFolders("inbox")
                .messages()
                .buildRequest()
                .select("from,isRead,receivedDateTime,subject")
                .top(2)
                .orderBy("receivedDateTime DESC")
                .get();
    }


    public static List<Message> getAllMessages() throws Exception {
        if (_userClient == null) {
            throw new Exception("Graph has not been initialized for user auth");
        }

        LinkedList<Option> requestOptions = new LinkedList<>();
        requestOptions.add(new HeaderOption("Prefer", "outlook.body-content-type=\"text\""));

        // Fetch messages from the inbox
        MessageCollectionPage messagePage = _userClient.me()
                .messages()
                .buildRequest(requestOptions)
                .select("subject,body,bodyPreview,uniqueBody")
                .get();

        return messagePage.getCurrentPage();
    }

    public static void saveInboxMessagesAsMIME() throws Exception {
        if (_userClient == null) {
            throw new Exception("Graph client not initialized");
        }

        // Start with the initial request  "inbox","SentItems"
        MessageCollectionRequest request = _userClient.me()
                .mailFolders("inbox")
                .messages()
                .buildRequest()
                .select("id,subject,receivedDateTime,from");

        int messageCount = 0;
        int folderNumber = 1; // Initialize folderNumber as 1

        while (request != null) {
            MessageCollectionPage page = request.get();
            for (var message : page.getCurrentPage()) {
                if (messageCount >= 100) {
                    // Reset the counter and increment the folder number
                    messageCount = 0;
                    folderNumber++;
                }
                String messageId = message.id;
                String messageSubject = message.subject.replace('/',' ').trim();
                int maxLength = 60;
                if (messageSubject.length() > maxLength)
                    messageSubject = messageSubject.substring(0, maxLength);
                messageSubject = new String(messageSubject.getBytes(),"UTF-8");
                //String messageFrom = message.from.toString();
                String messagereceivedDateTime = message.receivedDateTime.format(DateTimeFormatter.ofLocalizedDateTime(FormatStyle.SHORT)).replace('/','_');

                saveMessageAsMIME(messagereceivedDateTime,messageSubject,messageId, folderNumber);
                messageCount++;
            }
            // Get the next page request
            request = page.getNextPage() != null ? page.getNextPage().buildRequest() : null;
        }
    }

    private static void saveMessageAsMIME(String messagereceivedDateTime,String messageSubject,String messageId, int folderNumber) throws Exception {
        MessageStreamRequest streamRequest = _userClient.me().messages(messageId).content().buildRequest();

        Path folderPath = Paths.get("MIME_Messages_Folder" + folderNumber);
        Path mimeFilePath = folderPath.resolve(messagereceivedDateTime + "__" + messageSubject + ".eml");
        Files.createDirectories(folderPath);

        try (InputStream mimeStream = streamRequest.get();
             FileOutputStream fileOut = new FileOutputStream(mimeFilePath.toFile())) {
            mimeStream.transferTo(fileOut);
        }
    }
}



