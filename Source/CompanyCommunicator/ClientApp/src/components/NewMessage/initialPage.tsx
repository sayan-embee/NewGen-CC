// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { Header, Flex, Button, Card, CardBody, Text } from "@fluentui/react-northstar";

import { TFunction } from "i18next";

import { getBaseUrl } from '../../configVariables';

import './newMessage.scss';
import './teamTheme.scss';

// const image_Icon = getBaseUrl() + "/image/imageIcon.png";
// const pdf_Icon = getBaseUrl() + "/image/pdfIcon.png";
// const question_answer_Icon = getBaseUrl() + "/image/q&aIcon.png";
// const email_Icon = getBaseUrl() + "/image/emailIcon.png";



export interface IState {
    templateType?: any;
    selectedTemplate?: any;
}

export interface IProps extends WithTranslation {
    history?: any;
    location?: any
}

class InitialPageNewMessage extends React.Component<IProps, IState> {
    readonly localize: TFunction;

    constructor(props: IProps) {
        super(props);
        this.localize = this.props.t;

        this.state = {
            templateType: [
                { title: this.localize("ImageUpload"), image: getBaseUrl() + "/image/imageIcon.png" },
                { title: this.localize("PDFUpload"), image: getBaseUrl() + "/image/pdfIcon.png" },
                { title: this.localize("Q&AUpload"), image: getBaseUrl() + "/image/q&aIcon.png" },
                { title: this.localize("EmailUpload"), image: getBaseUrl() + "/image/emailIcon.png" },


            ]
        }

    }

    public async componentDidMount() {
        if (this.props.location.state) {
            this.setState({
                selectedTemplate: this.props.location.state.data
            })
        }

    }


    private onNext = (title: any) => {
        this.props.history.push({ pathname: "/newmessage", state: { data: title } })
    }

    public render(): JSX.Element {
        return (
            <div>
                 <Text styles={{display:"flex", justifyContent:"center"}} weight="bold" content={this.localize("ChooseTemplate")} />
                <Flex vAlign="center" hAlign="center" gap="gap.small" padding="padding.medium" className="templateTypeContainerFlex">
                    <div className="templateTypeContainer">
                   
                        {this.state.templateType.map((e: any) => {
                            return <Flex gap="gap.small" padding="padding.medium" className="templateTypeButtonFlex">
                                <Card className={`templateTypeButton ${(this.state.selectedTemplate === e.title) && 'selectedTemplateBackgroundColor'}`} size="medium" styles={{
                                    borderRadius: '6px',
                                    padding: '10px 10px 20px',
                                    backgroundColor: '#ffffff',
                                    boxShadow: '5px 5px 10px #e0e0e0',
                                    cursor: 'pointer',
                                    alignItems: 'center',
                                    ':hover': {
                                        backgroundColor: '#ffffff',
                                        boxShadow: '5px 5px 10px #e0e0e0',

                                    },
                                }} onClick={() => this.onNext(e.title)}>
                                    <CardBody></CardBody>
                                    <img src={e.image} className="imageIcon" />
                                    <Header as="h4" content={e.title} style={{ marginBottom: '0', color: "black" }} />
                                    {/* <Flex hAlign="end" vAlign="end" gap="gap.smaller">
                                    <Button className="curve-btn" title={this.localize("ImageUpload")} onClick={()=>this.viswas()} >  <img src={rightArrowImage}/>   </Button>
                                </Flex> */}
                                </Card>
                            </Flex>
                        })}
                    </div>

                </Flex>

            </div>

        )
    }




}

const newMessageWithTranslation = withTranslation()(InitialPageNewMessage);
export default newMessageWithTranslation;
