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


export const getInitAdaptiveCardPDFUpload = (t: TFunction) => {
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
                        "type": "Image",
                        "spacing": "Default",
                        "url": "",
                        "width": "80px",
                        "height": "80px",
                        "altText": "",
                        "horizontalAlignment": "left"
                    },
                    {
                        "type": "TextBlock",
                        "text": "",
                        "wrap": true
                    },
                    {
                        "type": "TextBlock",
                        "text": "",
                        "wrap": true
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

export const getCardTitlePDFUpload = (card: any) => {
    return card.body[0].text;
}

export const setCardTitlePDFUpload = (card: any, title: string) => {
    card.body[0].text = title;
}

export const getCardImageLinkPDFUpload = (card: any) => {
    return card.body[1].url;
}

export const setCardImageLinkPDFUpload = (card: any, imageLink?: string) => {
    card.body[1].url = imageLink;
}

export const getCardPdfNamePDFUpload = (card: any) => {
    return card.body[2].text;
}

export const setCardPdfNamePDFUpload = (card: any, link?: string) => {
    card.body[2].text = link;
}


export const getCardSummaryPDFUpload = (card: any) => {
    return card.body[3].text;
}

export const setCardSummaryPDFUpload = (card: any, summary?: string) => {
    card.body[3].text = summary;
}

export const getCardAuthorPDFUpload = (card: any) => {
    return card.body[4].text;
}

export const setCardAuthorPDFUpload = (card: any, author?: string) => {
    card.body[4].text = author;
}

export const getCardBtnTitlePDFUpload = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLinkPDFUpload = (card: any) => {
    return card.actions[0].url;
}

// set the values collection with buttons to the card actions
export const setCardBtnsPDFUpload = (card: any, values: any[]) => {
    if (values !== null) {
            card.actions = values;
    } else {
        delete card.actions;
    }
}


