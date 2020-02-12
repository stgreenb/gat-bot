# gat-bot

Base code for building WebEx Teams Bot hosted inside G suite using Google App Script (JavaScript).


## Business/Technical Challenge

We often hear that when creating simple WebEx teams bots, where to host the code and how to make it accessible can be slow down projects before they even get started. In today's world everybody has a google account. The goal is to show how easy it is to host the code inside of a google document. We'll show that google sheets can be a very workable database for your code along with hosting the application itself. As well, one of our peers created an awesome polling tool, but the tool died when he left Cisco, so why not rebuild it.

## Proposed Solution

We were inspired by [this article on building a Viber bot without code](https://developers.viber.com/blog/2017/09/12/build-a-bot-with-zero-coding) and want to take what [Itamar Mula](https://github.com/ItamarM) created and bring it to Webex Teams. For some this will be a simple poll bot they can use, but for others this will be the base of how to host a bot inside a g-suite document. Stretch goals include taking it further turning the first bot into a live polling tool that might be a simple version of what [Kahoot](https://getkahoot.com/) or [Poll Everywhere](https://www.polleverywhere.com/) do (using Google Slides). Now that buttons and cards are in Webex Teams, that's another way to respond(see image). 

![button image](https://github.com/stgreenb/gat-bot/blob/master/gat_with_buttons_and_card.jpg?raw=true)


### Products Technologies/ Services

Our solution will leverage the following technologies

* [Webex Teams](http://developer.webex.com)
* [Google App Script](https://developers.google.com/apps-script/)
* [G Suite](https://gsuite.google.com/)

## Team Members


* Steve Greenberg <stgreenb@cisco.com> - America's Partner Org
* Chris Norman <christno@cisco.com> - Enterprise



## Installation & Usage

1. **Get your own Copy.** Make a copy of [this spreadsheet](https://docs.google.com/spreadsheets/d/1ctIERb1yyptzXdIyQk-4RNt3cNvcHoNmH6HYrgzjhqQ/edit?usp=sharing). Now with Buttons & Cards! Once you open Google Sheets, Click File > Make a copy…

2. **Fill in your info in the userInput tab.** There are several tabs, but start with userInput. In order not to break anything, only change the yellow boxes. Start with your name and email. The other two boxes we'll get to next. 

3. **Get a bot token.** The top box, "Bot's Bearer Token" can be obtained by. 
   1. Visit https://developer.webex.com/my-apps/new and login with your webex teams account. 
   2. Select create a bot. 
   3. Fill in the required fields of Name, username, icon and description. The polls / surveys you send out will come from this bot, so think of a good name and icon to use. 
   4. Save the bots token back in the google sheet. _Be careful with the token. Anybody who has it can impersonate your bot._

4. **Get the url for your app script.**
   1. Open the Script editor... by clicking “Tools” > “Script editor...”
   2. Publish the sheet by clicking "Publish" > "Deploy as a Web App" 
   3. Select the latest project version to deploy.
   4. Select "me" (your account) as the "Execute as" dropdown.
   5. Select the "Anyone, even anonymous option" for the “Who has access to the app” dropdown
   6. Deploy the app and authorize the application.
   7. Copy the URL back into "Current web app URL" (you can close the script editor)

5. **Build recipients list on the recipients tab.** This can be done in 2 ways. 
   * Manually filling in or cutting and pasting the email addresses into column B "Webex Teams Emails". 
   * Have the users add themselves. From the "Form" menu you can forward the form or get a link to the form where users can add themselves. _NOTE: There is a current limitation that a user added in the middle of a poll will not get new questions._ 

6. **Build the questions tab.** For each question (currently Max 10 questions per poll) you wish to ask as part of the poll / survey fill in columns B through E. 
   * A. We said B through E, no need to touch A. 
   * B. The main question you are asking. Something like, "How would you rate your experience?"
   * C. A description of the possible answers. "A:Awesome, B:Neutral, C:Aweful" _Note: current answers can be EITHER letters or numbers. Fill in the blank will be added later._ 
   * D. The lowest number or letter that the user can enter. 
   * E. The highest number or letter that the user can enter. _Note: you will likely have poor results if this value is more than 100 over the "lowest" number._ 

7. **Start your poll.** A new menu "Poll" has been added to the sheet. Select "Start Poll". Once selected, each recipient will receive the fist question, and if there are subsequent questions, they will get them right after completing the previous.  Each question will go out sequentially to the users. 

8. **Watch the results.** Raw data of each response can be found on the "responses" tab and on the "dataStorage" tab you can find aggregate info as well as pre-created charts (you may have to scroll down to see all charts). All data and charts will live update as new responses come in. 

9. **End poll.** From the "Poll" menu select "Stop Poll". _Note: Once the poll has been stopped, the bot will no longer take any communication (its webhook is deleted). Also, for your next poll, you can skip right to step 5 (or 6 if your recipients are the same)._ 

  


## License

Provided under Cisco Sample Code License, for details see [LICENSE](./LICENSE.md)

## Code of Conduct

Our code of conduct is available [here](./CODE_OF_CONDUCT.md)

## Contributing

See our contributing guidelines [here](./CONTRIBUTING.md)
