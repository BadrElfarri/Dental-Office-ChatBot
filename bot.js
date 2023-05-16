// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, ActivityTypes } = require("botbuilder");
const { CustomQuestionAnswering } = require("botbuilder-ai");
const axios = require("axios");
const DentistScheduler = require("./dentistscheduler");

const scheduler = new DentistScheduler({
  SchedulerEndpoint: process.env.ScheduleEndpoint,
});

class CustomQABot extends ActivityHandler {
  constructor() {
    super();
    // If a new user is added to the conversation, send them a greeting message
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
          const DefaultWelcomeMessageFromConfig =
            process.env.DefaultWelcomeMessage;
          await context.sendActivity(
            DefaultWelcomeMessageFromConfig?.length > 0
              ? DefaultWelcomeMessageFromConfig
              : "Hello and Welcome"
          );
        }
      }

      await next();
    });

    // When a user sends a message, perform a call to the QnA Maker and LUIS service to retrieve best answer.
    this.onMessage(async (context, next) => {
      // If environment values are missing return a warning.
      if (
        !process.env.ProjectName ||
        !process.env.PredictionUrl ||
        !process.env.RequestId
      ) {
        const unconfiguredQnaMessage =
          "NOTE: \r\n" +
          "Custom Question Answering is not configured. To enable all capabilities, add `ProjectName`, `PredictionUrl`, `RequestId` and `LanguageEndpointHostName` to the .env file. \r\n" +
          "You may visit https://language.cognitive.azure.com/ to create a Orchestration Workflow combining a Custom Question Answering and Conversational Language Understanding Projects.";

        await context.sendActivity(unconfiguredQnaMessage);
      } else {
        console.log("Calling CQA");

        // Construct the request body with the user query
        const requestBody = {
          kind: "Conversation",
          analysisInput: {
            conversationItem: {
              id: context.activity.from.id,
              text: context.activity.text,
              modality: "text",
              participantId: context.activity.from.id,
            },
          },
          parameters: {
            projectName: process.env.ProjectName,
            verbose: true,
            deploymentName: "DentalOffice",
            stringIndexType: "TextElement_V8",
          },
        };
        try {
          // Make the HTTP POST request to the prediction URL
          const response = await axios.post(
            process.env.PredictionUrl,
            requestBody,
            {
              headers: {
                "Ocp-Apim-Subscription-Key": process.env.LanguageEndpointKey,
                "Apim-Request-Id": process.env.RequestId,
                "Content-Type": "application/json",
              },
            }
          );

          // console.log(response.data);
          // Process the response and send the appropriate message back to the user
          if (
            (response.data.result.prediction.topIntent === "LUIS") &
            (response.data.result.prediction.intents.LUIS.confidenceScore >=
              0.5)
          ) {
            if (
              response.data.result.prediction.intents.LUIS.result.prediction
                .topIntent === "ScheduleAppointment"
            ) {
              // Schedule a time for the patient
              const entities =
                response.data.result.prediction.intents.LUIS.result.prediction
                  .entities;
              if (entities.length > 0) {
                const scheduletime = entities.find(
                  (entity) => entity.category === "time"
                );
                const scheduleResult = await scheduler.scheduleAppointment(
                  scheduletime.text
                );
                await context.sendActivity(scheduleResult);
              } else {
                await context.sendActivity(
                  "Could not schedule a time, please provide a time in a form 8am or so"
                );
              }
            } else if (
              response.data.result.prediction.intents.LUIS.result.prediction
                .topIntent === "GetAvailability"
            ) {
              // Consume the availability endpoint
              const availability = await scheduler.getAvailability();
              console.log("Availability:");
              await context.sendActivity(
                "The available time slots are: " + availability
              );
            }
          } else if (
            (response.data.result.prediction.topIntent === "QNA") &
            (response.data.result.prediction.intents.QNA.confidenceScore >= 0.5)
          ) {
            const answer_qna =
              response.data.result.prediction.intents.QNA.result.answers[0]
                .answer;
            console.log("Answer QNA:", answer_qna);
            await context.sendActivity(answer_qna);
          } else {
            await context.sendActivity("No answers was found.");
          }
        } catch (error) {
          console.error("Error:", error);
          await context.sendActivity(
            "Failed to fetch an answer. error in communication"
          );
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }
}

module.exports.CustomQABot = CustomQABot;
