// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";
import * as AdaptiveCards from "adaptivecards";
import MarkdownIt from "markdown-it";

// Static method to render markdown on the adaptive card
AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    var md = new MarkdownIt();
    // Teams only supports a subset of markdown as per https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format?tabs=adaptive-md%2Cconnector-html#formatting-cards-with-markdown
    md.disable(['image', 'table', 'heading',
        'hr', 'code', 'reference',
        'lheading', 'html_block', 'fence',
        'blockquote', 'strikethrough']);
    // renders the text
    result.outputHtml = md.render(text);
    result.didProcess = true;
}


export const getInitAdaptiveCardEmailTemplate = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
        return (
            {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "text": titleTextAsString,
                        "size": "ExtraLarge",
                        "wrap": true
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "text": ""
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "text": ""
                    },
                    {
                        "type": "TextBlock",
                        "wrap": true,
                        "size": "Small",
                        "weight": "Lighter",
                        "text": ""
                    }
                ],
                "msteams": {
                    "width": "Full"
                },
                "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
            }
        );

}

export const getCardTitleEmailTemplate = (card: any) => {
    return card.body[0].text;
}

export const setCardTitleEmailTemplate = (card: any, title: string) => {
    card.body[0].text = title;
}


export const getCardFileNameEmailTemplate = (card: any) => {
    return card.body[1].text;
}

export const setCardFileNameEmailTemplate = (card: any, link?: string) => {
    card.body[1].text = link;
}

// export const getCardFileNametitleEmailTemplate = (card: any) => {
//     return card.actions[1].title;
// }

// export const setCardFileNameTitleEmailTemplate = (card: any, title?: string) => {
//     card.actions[1].title = title;
// }

export const getCardSummaryEmailTemplate = (card: any) => {
    return card.body[2].text;
}

export const setCardSummaryEmailTemplate = (card: any, summary?: string) => {
    card.body[2].text = summary;
}

export const getCardAuthorEmailTemplate = (card: any) => {
    return card.body[3].text;
}

export const setCardAuthorEmailTemplate = (card: any, author?: string) => {
    card.body[3].text = author;
}

export const getCardBtnTitleEmailTemplate = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLinkEmailTemplate = (card: any) => {
    return card.actions[0].url;
}

// set the values collection with buttons to the card actions
export const setCardBtnsEmailTemplate = (card: any, values: any[]) => {
    if (values !== null) {
            card.actions = values;
    } else {
        delete card.actions;
    }
}


