// Your OpenAI API key
// const OPENAI_API_KEY = ''; // Replace with your actual API key

/**
 * onOpen
 * ------
 * This function runs automatically when the Google Slides presentation is opened.
 * It performs two main tasks:
 * 1. Creates a custom menu item called "My Custom Tools" with the option "Generate Slide Deck"
 *    that calls the generateSlideDeck() function.
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('Generate Slide Deck')
    .addItem('Generate Slide Deck', 'generateSlideDeck')
    .addToUi();
}

/**
 * generateSlideDeck
 * -----------------
 * This function drives the creation of the slide deck based on a user-provided topic.
 * It follows these steps:
 * 1. Prompts the user to enter a topic for the slide deck.
 * 2. Prompts the user for an email ID to send the slide deck link.
 * 3. Uses the topic to call GPTForSlides(), which sends a crafted prompt to ChatGPT and returns
 *    formatted slide content.
 * 4. Retrieves the active presentation and deletes any existing slides.
 * 5. Splits the ChatGPT response into individual slides using the delimiter "$$$".
 * 6. For each slide, separates the title from the content using a colon (":") and then extracts
 *    bullet points enclosed in "< >".
 * 7. Creates a new slide for each set of content, setting the title and appending bullet points as
 *    individual paragraphs.
 * 8. Sends an email to the provided email ID with a link to the updated slide deck.
 */
function generateSlideDeck() {
  var ui = SlidesApp.getUi();

  // Prompt user for topic
  var topicResponse = ui.prompt("Enter the topic for your slide deck", ui.ButtonSet.OK_CANCEL);
  if (topicResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Action canceled.");
    return;
  }
  var topic = topicResponse.getResponseText().trim();
  if (!topic) {
    ui.alert("Topic cannot be empty.");
    return;
  }

  // Prompt user for email
  var emailResponse = ui.prompt("Enter the email ID to send the slide link", ui.ButtonSet.OK_CANCEL);
  if (emailResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Action canceled.");
    return;
  }
  var email = emailResponse.getResponseText().trim();
  if (!email || !email.includes("@")) {
    ui.alert("Please enter a valid email address.");
    return;
  }

  // Call ChatGPT to generate slide deck content based on the topic.
  var deckContent = GPTForSlides(topic, 500);
  if (deckContent.startsWith("Error:")) {
    ui.alert("Failed to generate slides: " + deckContent);
    return;
  }

  // Use the active presentation.
  var presentation = SlidesApp.getActivePresentation();

  // Delete all existing slides in the presentation.
  var slides = presentation.getSlides();
  slides.forEach(function(slide) {
    slide.remove();
  });

  // Split the response into individual slides based on the delimiter "$$$".
  var slidesData = deckContent.split('$$$');
  slidesData.forEach(function(slideData) {
    if (slideData.trim() === "") return;

    // Split by colon to separate the title from the content.
    var parts = slideData.split(':');
    if (parts.length < 2) return; // Ensure there's both a title and content.

    var title = parts[0].trim();
    var content = parts.slice(1).join(":").trim();

    // Extract bullet points from content, which are enclosed in "< >".
    var bulletPoints = content.match(/<([^>]+)>/g);
    if (!bulletPoints) return;

    // Create a new slide with a title and body layout.
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE).asShape().getText().setText(title);

    // Set the body content by appending each bullet point as a separate paragraph.
    var bodyText = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY).asShape().getText();
    bulletPoints.forEach(function(point) {
      // Remove the angle brackets and trim whitespace.
      var bulletPointText = point.replace(/[<>]/g, '').trim();
      // Prepend a bullet character to simulate bullet formatting.
      bodyText.appendParagraph('• ' + bulletPointText);
    });
  });

  // Get the link to the Google Slides presentation
  var presentationUrl = presentation.getUrl();

  // Send email with the link
  MailApp.sendEmail({
    to: email,
    subject: "Slide Deck: " + topic,
    body: "Your slide deck on the topic '" + topic + "' has been generated. You can view it here:\n\n" + presentationUrl
  });

  ui.alert("Slide deck updated! A link has been sent to " + email);
}

/**
 * GPTForSlides
 * ------------
 * This function communicates with the OpenAI API to generate slide deck content.
 */
function GPTForSlides(topic, maxTokens) {
  if (!topic) return "Error: No topic provided.";

  var apiUrl = 'https://api.openai.com/v1/chat/completions';

  var payload = {
    'model': 'gpt-4o-mini', // or whichever model you're using
    'messages': [
      {
        'role': 'system',
        'content': 'You are an expert in creating Google Slides content. Provide slide titles and bullet point content for a slide deck on the given topic. For each slide, provide a relevant title and bullet point content, formatted as follows: ' +
        '[Slide Title]: <bullet point sentence 1> <bullet point sentence 2> <bullet point sentence 3> $$$ \n\n' +
        'Important Instructions:\n' +
        '- Separate each slide by a $$$.\n' +
        '- Ensure that titles and bullet points are clearly separated by a colon (:) and bullet points are surrounded by angle brackets (< >).\n' +
        '- Do not use the literal text "Slide Title"; instead, provide a unique title that reflects the slide content using plain text only.\n' +
        '- Make the bullet points complete sentences with a positive and educational tone about the selected topic.\n' +
        '- Do not add any extra text or formatting beyond the specified format.'
      },
      {'role': 'user', 'content': "Generate a slide deck for the topic: " + topic}
    ],
    'max_tokens': maxTokens
  };
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {'Authorization': 'Bearer ' + OPENAI_API_KEY},
    'payload': JSON.stringify(payload)
  };
  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var json = JSON.parse(response.getContentText());
    return json.choices[0].message.content.trim();
  } catch (error) {
    return "Error: " + error.toString();
  }
}

