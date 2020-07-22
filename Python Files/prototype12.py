import os
import sys
import datetime
from datetime import timedelta
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sb
import re
import glob
import csv
from urllib.parse import urljoin
from flask import Flask, request, Response
from webexteamssdk import WebexTeamsAPI, Webhook
from PIL import Image, ImageDraw, ImageFont
import json

# Script metadata
__author__ = "JP Shipherd"
__author_email__ = "jshipher@cisco.com"
__contributors__ = ["Chris Lunsford <chrlunsf@cisco.com>"]
__copyright__ = "Copyright (c) 2016-2020 Cisco and/or its affiliates."
__license__ = "MIT"


# Constants
WEBHOOK_NAME = "botWithCardExampleWebhook"
WEBHOOK_URL_SUFFIX = "/events"
MESSAGE_WEBHOOK_RESOURCE = "messages"
MESSAGE_WEBHOOK_EVENT = "created"
CARDS_WEBHOOK_RESOURCE = "attachmentActions"
CARDS_WEBHOOK_EVENT = "created"
MEMBERSHIP_WEBHOOK_RESOURCE = "memberships"
MEMBERSHIP_WEBHOOK_EVENT="created"

# Adaptive Card Design Schema for a sample form.
# To learn more about designing and working with buttons and cards,
# checkout https://developer.webex.com/docs/api/guides/cards
CARD_DATECOMPARE={
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.1",
    "body": [
        {
            "type": "TextBlock",
            "text": "Date Comparison Feature:",
            "size": "medium",
            "weight": "bolder",
            "wrap": True
        },
        {
            "type": "TextBlock",
            "text": "The **Date Comparison** feature allows you to enter different parameters of your liking, as well as a start and end date. You will then see a comparison between the selected parameters on those dates in both values and satellite images."
        },
        {
            "type": "TextBlock",
            "text": "Please select your country of choice",
            "wrap": True
        },
        {
            "type": "Input.ChoiceSet",
            "id": "CountryChoiceVal",
            "value": "Japan",
            "choices":[
                {
                    "title": "Japan",
                    "value": "Japan"
                }]
        },
        {
            "type": "TextBlock",
            "text": "Let's get a comparison of your start/end date. 				Please enter your **start date**:",
            "wrap": True
        },
        {
            "type": "Input.Date",
            "id": "start_date",
        },
        {
            "type": "TextBlock",
            "text": "Please enter your **end date**:",
            "wrap": True
        },
        {
            "type": "Input.Date",
            "id": "end_date"
        },
        {
            "type": "TextBlock",
            "text": "Which **parameter** do you want to include?",
            "wrap": True
        },

        {
            "type": "Input.Toggle",
            "title": "Pollution",
            "id": "Pollution",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Vegetation Index",
            "id": "Forest Loss",
            "wrap": True,
            "value": "false"
        },
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Submit",
            "data": {
                "formDemoAction": "SubmitA"
            }
        }
    ],
    }

CARD_INSIGHTS={
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.1",
    "body": [
        {
            "type": "TextBlock",
            "text": "Pandemic Insights Feature:",
            "size": "medium",
            "weight": "bolder",
            "wrap": True
        },        
        {   
            "type": "TextBlock",
            "text": "The **Pandemic Insights** feature allows you to enter different parameters of your liking. You will then be able to see the correlation of your chosen parameters with the COVID-19 situation in your chosen country",
            "wrap": True
        },
        {
            "type": "TextBlock",
            "text": "Please select your country of choice",
            "wrap": True
        },
        {
            "type": "Input.ChoiceSet",
            "id": "CountryChoiceVal",
            "value": "Japan",
            "choices":[
                {
                    "title": "Japan",
                    "value": "Japan"
                }]
        },
        {   
            "type": "TextBlock",
            "text": "Let's gain more insights on your selected country during the COVID-19 pandemic. Which **parameter** do you want to include?",
            "wrap": True
        },
        {
            "type": "Input.Toggle",
            "title": "COVID-19 Infections",
            "id": "Infections",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "COVID-19 Deaths",
            "id": "Deaths",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Population",
            "id": "Population",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Gross Domestic Product (GDP)",
            "id": "GDP",
            "wrap": True,
            "value": "false"
        },
                {
            "type": "Input.Toggle",
            "title": "Healthcare per Capita",
            "id": "Healthcare",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Pollution Levels",
            "id": "Pollution",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Forest per Land ratio",
            "id": "Forest ratio",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Average Temperature",
            "id": "Temperature",
            "wrap": True,
            "value": "false"
        },
        {
            "type": "Input.Toggle",
            "title": "Humidity",
            "id": "Humidity %",
            "wrap": True,
            "value": "false"
        },
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Submit",
            "data": {
                "formDemoAction": "SubmitB"
            }
        }
    ]
  }

CARD_FORECAST={
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.1",
    "body": [
        {
            "type": "TextBlock",
            "text": "Parameter Forecast Feature:",
            "size": "medium",
            "weight": "bolder",
            "wrap": True
        },        
        {   
            "type": "TextBlock",
            "text": "The **Parameter Forecast** feature allows you to enter a parameter of your liking, as well as a date. You will then be able to see the forecasted value of your chosen parameter on the selected date together with a 95% confidence interval for your chosen location.",
            "wrap": True
        },
        {
            "type": "TextBlock",
            "text": "Please select your location of choice",
            "wrap": True
        },
        {
            "type": "Input.ChoiceSet",
            "id": "CountryChoiceVal",
            "value": "Tokyo",
            "choices":[
                {
                    "title": "Tokyo",
                    "value": "Tokyo"
                }]
        },
        {   
            "type": "TextBlock",
            "text": "Let's get a forecasted value on your selected location. Which **date** are you interested in? Please ensure dates selected are in the **future**. Do note that dates too far in the future will not have a forecasted value.",
            "wrap": True
        },
        {
            "type": "Input.Date",
            "id": "forecast_date"
        },
        {   
            "type": "TextBlock",
            "text": "Which **parameter** are you interested in?",
            "wrap": True
        },
        {
            "type": "Input.ChoiceSet",
            "id": "ParamChoiceVal",
            "value": "NO2 Pollution (Weekly Forecast)",
            "choices":[
                {
                    "title": "NO2 Pollution (Weekly Forecast)",
                    "value": "Pollution"
                },
                {
                    "title": "Vegetation Index (Monthly Forecast)",
                    "value": "NDVI"
                }
                ]
        },
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Submit",
            "data": {
                "formDemoAction": "SubmitC"
            }
        }
    ]
  }
  
compare_response = {
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [    	
    	{
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "About Japan",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "Japan is a island country in East Asia located in the Pacific Ocean. It comprises an archipelago of 6,852 islands covering 377,975 square kilometers (145,937 sq mi); its five main islands, from north to south, are Hokkaido, Honshu, Shikoku, Kyushu, and Okinawa. It's largest city, Tokyo, is also the country's capital.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "Japan is widely considered to be a leading example of a first world nation renowned for its technology and rich culture. With the third highest GDP in the world, it also ranks very highly in the Human Development Index, boasting the second highest life expectancy. Some major issues the country faces include its aging population and disaster-prone geography.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "weight": "Bolder",
            "text": "Date Comparison response",
            "size": "medium",
            "wrap": True
        },        
        {
            "type": "TextBlock",
            "text": "The values of your selected parameters as well as satellite images based on your selected dates are available for your viewing below. Should you not select a parameter or request values from an invalid date, no results will be shown.",
            "wrap": True
        },
        ],
        "actions":[
        {
            "type": "Action.ShowCard",
            "title": "Pollution Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [
        {
              "type": "TextBlock",
              "text": "__TEXT1__"
        },   
        {
              "type": "Image",
              "url": "__DATA1__",
              "size": "Medium",
              "height":"200px"
        }, 
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},      
    	{
            "type": "Action.ShowCard",
            "title": "Vegetation Index Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [   
        {
              "type": "TextBlock",
              "text": "__TEXT2__"
        },   
        {
              "type": "Image",
              "url": "__DATA2__",
              "size": "Medium",
              "height":"190px"
        }, 
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},
    ],    
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
}

compare_response2 = {
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [    	
    	{
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "About Japan",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "Japan is a island country in East Asia located in the Pacific Ocean. It comprises an archipelago of 6,852 islands covering 377,975 square kilometers (145,937 sq mi); its five main islands, from north to south, are Hokkaido, Honshu, Shikoku, Kyushu, and Okinawa. It's largest city, Tokyo, is also the country's capital.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "Japan is widely considered to be a leading example of a first world nation renowned for its technology and rich culture. With the third highest GDP in the world, it also ranks very highly in the Human Development Index, boasting the second highest life expectancy. Some major issues the country faces include its aging population and disaster-prone geography.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "weight": "Bolder",
            "text": "Date Comparison response",
            "size": "medium",
            "wrap": True
        },        
        {
            "type": "TextBlock",
            "text": "The values of your selected parameters as well as satellite images based on your selected dates are available for your viewing below. Should you not select a parameter or request values from an invalid date, no results will be shown.",
            "wrap": True
        },
        ],
        "actions":[
        {
            "type": "Action.ShowCard",
            "title": "Pollution Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [
        {
              "type": "TextBlock",
              "text": "__TEXT1__"
        },   
        {
              "type": "Image",
              "url": "__DATA1__",
              "size": "Medium",
              "height":"200px"
        }, 
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},      
    	{
            "type": "Action.ShowCard",
            "title": "Vegetation Index Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [   
        {
              "type": "TextBlock",
              "text": "This parameter was not selected for comparison. Please select the Vegetation Index toggle if you are interested in this parameter."
        }],   
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},
    ],    
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
}

compare_response3 = {
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [    	
    	{
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "About Japan",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "Japan is a island country in East Asia located in the Pacific Ocean. It comprises an archipelago of 6,852 islands covering 377,975 square kilometers (145,937 sq mi); its five main islands, from north to south, are Hokkaido, Honshu, Shikoku, Kyushu, and Okinawa. It's largest city, Tokyo, is also the country's capital.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "Japan is widely considered to be a leading example of a first world nation renowned for its technology and rich culture. With the third highest GDP in the world, it also ranks very highly in the Human Development Index, boasting the second highest life expectancy. Some major issues the country faces include its aging population and disaster-prone geography.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "weight": "Bolder",
            "text": "Date Comparison response",
            "size": "medium",
            "wrap": True
        },        
        {
            "type": "TextBlock",
            "text": "The values of your selected parameters as well as satellite images based on your selected dates are available for your viewing below. Should you not select a parameter or request values from an invalid date, no results will be shown.",
            "wrap": True
        },
        ],
        "actions":[
        {
            "type": "Action.ShowCard",
            "title": "Pollution Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [
        {
              "type": "TextBlock",
              "text": "This parameter was not selected for comparison. Please select the Pollution toggle if you are interested in this parameter."
        },   
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},      
    	{
            "type": "Action.ShowCard",
            "title": "Vegetation Index Comparison",
            "card":{
                "type": "AdaptiveCard",
                "body": [   
        {
              "type": "TextBlock",
              "text": "__TEXT1__"
        },   
        {
              "type": "Image",
              "url": "__DATA1__",
              "size": "Medium",
              "height":"190px"
        }, 
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},
    ],    
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
}

insights_response= {
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [        
    	{
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "About Japan",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "Japan is a island country in East Asia located in the Pacific Ocean. It comprises an archipelago of 6,852 islands covering 377,975 square kilometers (145,937 sq mi); its five main islands, from north to south, are Hokkaido, Honshu, Shikoku, Kyushu, and Okinawa. It's largest city, Tokyo, is also the country's capital.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "Japan is widely considered to be a leading example of a first world nation renowned for its technology and rich culture. With the third highest GDP in the world, it also ranks very highly in the Human Development Index, boasting the second highest life expectancy. Some major issues the country faces include its aging population and disaster-prone geography.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "Pandemic Insights response",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "A **Correlation Table** as well as a **Pairplot Graph** based on your selected parameters are available for viewing. Some explanation will also be given to allow for your better understanding of each image shown.",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "Both images are derived using data gathered on a prefectural basis. Pollution data was collected using Google Earth's Engine Tropomi Explorer, while the other parameters were collected from National census data conducted by the Japanese Government.",
            "wrap": True
        }, 
        ],
        "actions":[
        {
            "type": "Action.ShowCard",
            "title": "Correlation Table",
            "card":{
                "type": "AdaptiveCard",
                "body": [
        {
              "type": "Image",
              "url": "__DATA1__",
              "size": "Medium",
              "height":"450px"
        },
        {
              "type": "TextBlock",
              "text": "__TEXT1__"
        },
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},
        {
            "type": "Action.ShowCard",
            "title": "Pairplot Graph",
            "card":{
                "type": "AdaptiveCard",
                "body":[ 
        {
              "type": "Image",
              "url": "__DATA2__",
              "size": "Medium",
              "height":"350px"
        },
        {
              "type": "TextBlock",
              "text": "__TEXT2__"
        },
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},  
        ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
}

forecast_response= {
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [        
        {
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "Parameter Forecast response",
            "wrap": True
        },         
        {
            "type": "TextBlock",
            "text": "__TEXT1__",
            "wrap": True
        },
        {
            "type": "TextBlock",
            "text": "__TEXT1.5__",
            "wrap": True
        },  
        {
            "type": "TextBlock",
            "text": "For the full time series forecast, please click on the _Full Forecast_ button below",
            "wrap": True
        }, 
        {
            "type": "TextBlock",
            "text": "For technical details behind the forecasted values, please click on the _More Information_ button below",
            "wrap": True
        }, 
        ],
        "actions":[
        {
            "type": "Action.ShowCard",
            "title": "Full Forecast",
            "card":{
                "type": "AdaptiveCard",
                "body": [
        {
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "Full Forecast Timeseries",
            "wrap": True
        },    
        {
              "type": "Image",
              "url": "__DATA1__",
              "size": "Medium",
              "height":"195px"
        },
        {
              "type": "TextBlock",
              "text": "__TEXT2__"
        },
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},
        {
            "type": "Action.ShowCard",
            "title": "More Information",
            "card":{
                "type": "AdaptiveCard",
                "body":[                 
        {
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "Forecasting Model Used",
            "wrap": True
        },    
        {
              "type": "TextBlock",
              "text": "__TEXT3__"
        },
        {
            "type": "TextBlock",
            "weight":"Bolder",
            "size": "medium",
            "text": "Error in Model",
            "wrap": True
        }, 
        {
              "type": "TextBlock",
              "text": "__TEXT4__"
        },
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        }},  
        ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
}

# Module variables
webhook_url = os.environ.get("WEBHOOK_URL", "")
port = int(os.environ.get("PORT", 0))
access_token = os.environ.get("WEBEX_TEAMS_ACCESS_TOKEN", "")
if not all((webhook_url, port, access_token)):
    print(
        """Missing required environment variable.  You must set:
        * WEBHOOK_URL -- URL for Webex Webhooks (ie: https://2fXX9c.ngrok.io)
        * PORT - Port for Webhook URL (ie: the port param passed to ngrok)
        * WEBEX_TEAMS_ACCESS_TOKEN -- Access token for a Webex bot
        """
    )
    sys.exit()

# Initialize the environment
# Create the web application instance
flask_app = Flask(__name__)
# Create the Webex Teams API connection object
api = WebexTeamsAPI()
# Get the details for the account who's access token we are using
me = api.people.me()


# Helper functions
def delete_webhooks_with_name():
    """List all webhooks and delete webhooks created by this script."""
    for webhook in api.webhooks.list():
        if webhook.name == WEBHOOK_NAME:
            print("Deleting Webhook:", webhook.name, webhook.targetUrl)
            api.webhooks.delete(webhook.id)


def create_webhooks(webhook_url):
    """Create the Webex Teams webhooks we need for our bot."""
    print("Creating Message Created Webhook...")
    webhook = api.webhooks.create(
        resource=MESSAGE_WEBHOOK_RESOURCE,
        event=MESSAGE_WEBHOOK_EVENT,
        name=WEBHOOK_NAME,
        targetUrl=urljoin(webhook_url, WEBHOOK_URL_SUFFIX)
    )
    print(webhook)
    print("Webhook successfully created.")

    print("Creating Attachment Actions Webhook...")
    webhook = api.webhooks.create(
        resource=CARDS_WEBHOOK_RESOURCE,
        event=CARDS_WEBHOOK_EVENT,
        name=WEBHOOK_NAME,
        targetUrl=urljoin(webhook_url, WEBHOOK_URL_SUFFIX)
    )
    print(webhook)
    print("Webhook successfully created.")
    
    print("Creating Membership Webhook...")
    webhook = api.webhooks.create(
        resource=MEMBERSHIP_WEBHOOK_RESOURCE,
        event=MEMBERSHIP_WEBHOOK_EVENT,
        name=WEBHOOK_NAME,
        targetUrl=urljoin(webhook_url, WEBHOOK_URL_SUFFIX)
    )
    print(webhook)
    print("Webhook successfully created.")

shared_dir = os.getcwd() + "/shared"

# for debug
import inspect
def debug_object(obj, prefix="==>"):
    print(f"{prefix} {type(obj)}")
    print(f"{str(obj)}")
    maxlen = max([len(i) for i in dir(obj) if not i.startswith("_")])
    for i in dir(obj):
        if not i.startswith("_"):
            attr = eval(f"obj.{i}")
            t = str(type(attr))
            print(f"  {i.ljust(maxlen)}: {t}")
            if "str" in t:
                print(f"    {attr}")
            elif "dict" in t:
                for k,v in attr.items():
                    print(f"    {k}: {v}")
            elif "method" in t:
                print(f"    {inspect.getfullargspec(attr)}")
            else:
                pass

import base64
def encode_local_data(filename, ctype="image/png"):
    # data URL is defined in RFC 2397
    # the value should be formed;
    #     "data:[<media type>][;base64],<data>"
    return "data:{};base64,{}".format(ctype, base64.b64encode(
            open(filename,"rb").read()).decode("utf-8"))

import copy

def putval(content, key, val, target="__DATA1__", encode=False, ctype=None,
           stored=True):
    """replace into val from target of key in content.
    return content replaced if sucessful, or return None if faild.
    note that it modifies content directly.
    you have to store it before calling this function.
    """
    if stored:
        _c = copy.deepcopy(content)
    else:
        _c = content
    def lookup_dict(d):
        for k,v in d.items():
            if k == key and v == target:
                if encode:
                    d.update({ k: encode_local_data(val, ctype) })
                else:
                    d.update({ k: val })
                return True
            elif isinstance(v, dict):
                if lookup_dict(v):
                    return True
            elif isinstance(v, list):
                for j in v:
                    if lookup_dict(j):
                        return True
        return False
    #
    ret = lookup_dict(_c)
    if ret:
        return _c
    else:
        return None

def putvals(content, vals, stored=True):
    """replace into each val from target of key in content.
    valus must a list of below items:
        {
            "key":key,
            "val":val,
            "target":"__DATA1__",
            "ctype":None,
            "ctype":ctype
        }
    """
    if stored:
        _c = copy.deepcopy(content)
    else:
        _c = content
    for x in vals:
        _c = putval(_c, x["key"], x["val"], target=x.get("target"),
                     encode=x.get("encode"), ctype=x.get("ctype"))
        if _c is None:
            return None
    return _c

def respond_to_membership(webhook):
    """Respond to new members"""

    # Some server side debugging
    room = api.rooms.get(webhook.data.roomId)
    members = api.memberships.get(webhook.data.id)
    person = api.people.get(members.personId)
    print(
        f"""
        NEW MEMBER IN ROOM '{room.title}'
        FROM '{person.displayName}'
        """
    )
    api.messages.create(
           room.id,
           markdown=f"**Hi there! I'm a COVID-19 information bot. Here are my available functions:**\n"    	                 
                    f"\n`@Bot /help:` Details of how to use the bot\n"
                    f"\n`@Bot /features:` Details on what functionalities the bot has\n"
                    f"\n`@Bot /comparision:` Activates the date comparison mode\n"  
                    f"\n`@Bot /insights:` Activates the pandemic insights mode\n"
                    f"\n`@Bot /forecast:` Activates the parameter forecast mode\n"
                    f"\n **Note:** If you're using a 1:1 space with the bot, please exclude `@Bot` when calling functions." 
                   )

def pollutionJP_datecompare(start_date, end_date):    
    try:
        with open(r'/home/shawn/Downloads/ee-chart (3).csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
        pan_mean=float(data[1][3])
        pan_mean_2dp=round(pan_mean,2)
        bef_mean=float(data[1][4])
        bef_mean_2dp=round(bef_mean,2)
        #dicts used for date formatting    
        calendar={"01":"January", "02":"February", "03":"March", "04":"April", "05":"May", "06":"June", "07":"July", "08":"August", "09":"September", "10":"October", "11":"November", "12":"December"}
        calendar_abr={"01":"Jan", "02":"Feb", "03":"Mar", "04":"Apr", "05":"May", "06":"Jun", "07":"Jul", "08":"Aug", "09":"Sep", "10":"Oct", "11":"Nov", "12":"Dec"} 
        calendar_no={"Jan":1, "Feb":2, "Mar":3, "Apr":4, "May":5, "Jun":6, "Jul":7, "Aug":8, "Sep":9, "Oct":10, "Nov":11, "Dec":12}
        
        #format starting date
        start_datelist=start_date.split("-")
        start_year=str(start_datelist[0])
        start_day=str(start_datelist[2])
        start_month=calendar[start_datelist[1]]
        start_month_abr=calendar_abr[start_datelist[1]]
        startdate=start_day+" "+str(start_month)+", "+start_year
        d1=datetime.datetime(int(start_year),int(calendar_no[calendar_abr[start_datelist[1]]]),int(start_day))
        
        #format ending date
        end_datelist=end_date.split("-")
        end_year=str(end_datelist[0])
        end_day=str(end_datelist[2])
        end_month=calendar[end_datelist[1]]
        end_month_abr=calendar_abr[end_datelist[1]]
        enddate=end_day+" "+str(end_month)+", "+end_year
        d2=datetime.datetime(int(end_year),int(calendar_no[calendar_abr[end_datelist[1]]]),int(end_day))
        
        #ensure date inputs are valid
        if d1>d2:
            return "Please ensure start date is before end date."
        elif d1<=datetime.datetime.now() and d1>=datetime.datetime(2018,7,11) and d2<=datetime.datetime.now() and d2>=datetime.datetime(2018,7,11):
            #get start date pollution
            start_date_object = datetime.datetime.strptime(startdate, "%d %B, %Y").timetuple().tm_yday
            start_year_no=int(start_year)-2018
            start_date_value=start_year_no*365+int(start_date_object)
            #allow dates between 9 day intervals to have values
            start_n= ((start_date_value-192)//9)
            if data[start_n+1][10] != "":
                start_pollution=str(data[start_n+1][10])
            #handle case where a 9-day value is absent
            elif data[start_n+1][10] == "":
                start_counter=0
                while data[start_n+1+start_counter][10] is "":
                    start_counter=start_counter-1
                start_pollution=str(data[start_n+1+start_counter][10])
#get end date pollution
            end_date_object = datetime.datetime.strptime(enddate, "%d %B, %Y").timetuple().tm_yday
            end_year_no=int(end_year)-2018
            end_date_value=end_year_no*365+int(end_date_object)
            #allow dates between 9 day intervals to have values
            end_n= ((end_date_value-192)//9)
            if data[end_n+1][10] != "":
                end_pollution=str(data[end_n+1][10])
            #handle case where a 9-day value is absent
            elif data[end_n+1][10] == "":
                end_counter=0
                while data[end_n+1+end_counter][10] is "":
                    end_counter=end_counter-1
                end_pollution=str(data[end_n+1+end_counter][10])
            
            if float(start_pollution)>float(end_pollution):
                change="a decrease"
                change_value= (float(start_pollution)-float(end_pollution))/float(end_pollution)
                percent_change=round(change_value*100,2)
            elif float(start_pollution)<float(end_pollution):
                change="an increase"
                change_value=(float(end_pollution)-float(start_pollution))/float(start_pollution)
                percent_change=round(change_value*100,2)
            else:
                return "The pollution level on both "+startdate+" and "+enddate+" is **"+start_pollution+" µmol/m2**. You can visualise this from the satellite images below."
            message="The pollution level on "+ startdate +" is **" + start_pollution +" µmol/m2** compared to **"+end_pollution+ " µmol/m2** on "+ enddate +\
            ". This represents **"+ str(change) + "** of **"+ str(percent_change)+"%**. For reference, the mean pollution levels in Japan during the pandemic (24 January, 2020 to present) is **"+ str(pan_mean_2dp)+ " µmol/m2** while the pollution levels before the pandemic is **"+ str(bef_mean_2dp)+ "µmol/m2**. You can visualise the difference from the satellite images below." 
            return message 

        elif d1>=datetime.datetime.now() or d2>=datetime.datetime.now():
            return "Please make sure the date inputs are not in the future."
        elif d1<datetime.datetime(2018,7,11) or d2<datetime.datetime(2018,7,11):
            return "Pollution levels are not available before 20 July, 2018."
        
    except:
        return "Please ensure that the date inputs are not in the future. Otherwise, information on the requested date is not available"

def append_images(images, direction='horizontal',
                  bg_color=(255,255,255), aligment='center'):
    """
    Appends images in horizontal/vertical direction.

    Args:
        images: List of PIL images
        direction: direction of concatenation, 'horizontal' or 'vertical'
        bg_color: Background color (default: white)
        aligment: alignment mode if images need padding;
           'left', 'right', 'top', 'bottom', or 'center'

    Returns:
        Concatenated image as a new PIL image object.
    """
    widths, heights = zip(*(i.size for i in images))

    if direction=='horizontal':
        new_width = sum(widths)
        new_height = max(heights)
    else:
        new_width = max(widths)
        new_height = sum(heights)

    new_im = Image.new('RGB', (new_width, new_height), color=bg_color)


    offset = 0
    for im in images:
        if direction=='horizontal':
            y = 0
            if aligment == 'center':
                y = int((new_height - im.size[1])/2)
            elif aligment == 'bottom':
                y = new_height - im.size[1]
            new_im.paste(im, (offset, y))
            offset += im.size[0]
        else:
            x = 0
            if aligment == 'center':
                x = int((new_width - im.size[0])/2)
            elif aligment == 'right':
                x = new_width - im.size[0]
            new_im.paste(im, (x, offset))
            offset += im.size[1]
    return new_im

def date_select(date):    
    try:
        with open(r'/home/shawn/Downloads/ee-chart (3).csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))

        #dicts used for date formatting    
        calendar={"01":"January", "02":"February", "03":"March", "04":"April", "05":"May", "06":"June", "07":"July", "08":"August", "09":"September", "10":"October", "11":"November", "12":"December"}
        calendar_abr={"01":"Jan", "02":"Feb", "03":"Mar", "04":"Apr", "05":"May", "06":"Jun", "07":"Jul", "08":"Aug", "09":"Sep", "10":"Oct", "11":"Nov", "12":"Dec"} 
        calendar_no={"Jan":1, "Feb":2, "Mar":3, "Apr":4, "May":5, "Jun":6, "Jul":7, "Aug":8, "Sep":9, "Oct":10, "Nov":11, "Dec":12}
        
        #format starting date
        datelist=date.split("-")
        year=str(datelist[0])
        day=str(datelist[2])
        month=calendar[datelist[1]]
        month_abr=calendar_abr[datelist[1]]
        date_full=day+" "+str(month)+", "+year
        d1=datetime.datetime(int(year),int(calendar_no[calendar_abr[datelist[1]]]),int(day))
        
        if d1<=datetime.datetime.now() and d1>=datetime.datetime(2018,7,11):
            #get date pollution
            date_object = datetime.datetime.strptime(date_full, "%d %B, %Y").timetuple().tm_yday
            year_no=int(year)-2018
            date_value=year_no*365+int(date_object)
            #allow dates between 9 day intervals to have values
            n= ((date_value-192)//9)
            if data[n+1][8] is not '':
                date_output=str(data[n+1][8])
                return date_output
            elif data[n+1][8] is '':
                counter=0
                while data[n+1+counter][8] is '':
                    counter=counter-1
                date_output=str(data[n+1+counter][8])
                return date_output
        elif d1>=datetime.datetime.now() or d1<datetime.datetime(2018,7,11):
            return "null"
        
            
    except IndexError:
        return "null"

def image_output(start_date, end_date):
    image_dir=r'/home/shawn/Sat_Images'
    
    start_img=image_dir+'/'+str(date_select(start_date))+'.png'
    end_img=image_dir+'/'+str(date_select(end_date))+'.png'
    req_img=[start_img,end_img]
    all_img=[]
    img_path=[]
    for file in glob.iglob(image_dir + '**/*'):
        all_img.append("".join(file))
    
    for img in req_img:
        if img in all_img:
            img_path.append(img)
    if len(img_path)<2 or len(pollutionJP_datecompare(start_date, end_date))<120:
        return "null"
    
    elif len(img_path)==2:
        #start image
        im = Image.open(img_path[0]) 
        draw = ImageDraw.Draw(im)
        font = ImageFont.truetype("arial.ttf", 25)
        (x, y) = (1050, 750)
        message = str(start_date)
        color = 'rgb(255, 255, 255)' # white color
        # draw the message on the background
        draw.text((x, y), message, fill=color, font=font)
        cropped = im.crop((725, 250, 1200, 800))
        
        #separator image
        im1_5=Image.open(r"/home/shawn/Sat_Images/split.png")
        cropped1_5 = im.crop((518, 400, 528, 950))
        
        #end image
        im2 = Image.open(img_path[1])
        draw = ImageDraw.Draw(im2)
        font = ImageFont.truetype("arial.ttf", 25)
        (x, y) = (1050, 750)
        message = str(end_date)
        color = 'rgb(255, 255, 255)' # white color
        # draw the message on the background
        draw.text((x, y), message, fill=color, font=font)
        cropped2 = im2.crop((725, 250, 1200, 800))
        
        #sidebar image
        im3= Image.open(r"/home/shawn/Sat_Images/sidebar.png")
        cropped3= im3.crop((0, 0, 160, 550))
        
        x=append_images([cropped,cropped1_5,cropped2,cropped3], direction='horizontal')
        x.save("/home/shawn/shared/pol_comparison.png")

def NDVI_value(date_input):    
    try:
        dt = datetime.datetime.strptime(date_input,'%Y-%m-%d').date()
        d=dt.strftime('%Y-%m-%d')
        year_no=int(d[0:4])-2018
        date_object = int(dt.timetuple().tm_yday)
        n= ((date_object-1)//16)
        with open(r'/home/shawn/Downloads/NDVI.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
        if dt<=datetime.date.today() and dt>=datetime.datetime.strptime('01/01/2018','%m/%d/%Y').date():
            if dt<=datetime.datetime.strptime(data[58][0],'%Y/%m/%d').date():
                return data[n+1+year_no*23][1]
            else:
                return data[58][1]
        elif dt>datetime.date.today() or dt<datetime.datetime.strptime('01/01/2018','%m/%d/%Y').date():
            return "null"
    except ValueError:
        return "null"
def NDVI_compare(start, end):
    try:
        start_veg=float(NDVI_value(start))
        end_veg=float(NDVI_value(end))
        with open(r'/home/shawn/Downloads/NDVI.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
        prior_veg=round(float(data[1][2]),3)
        during_veg=round(float(data[1][3]),3)
        
        start_date=datetime.datetime.strptime(start,'%Y-%m-%d').date().strftime('%d %B, %Y')
        end_date=datetime.datetime.strptime(end,'%Y-%m-%d').date().strftime('%d %B, %Y')
        #if start date after or equal end date
        if datetime.datetime.strptime(start,'%Y-%m-%d').date()>=datetime.datetime.strptime(end,'%Y-%m-%d').date():
            print("error")
            return "Please ensure start date is before end date."
        #if start date before end date
        else:
            if float(start_veg)>float(end_veg):
                change="a decrease"
                change_value= (float(start_veg)-float(end_veg))/float(end_veg)
                percent_change=round(change_value*100,2)
            elif float(start_veg)<float(end_veg):
                change="an increase"
                change_value=(float(end_veg)-float(start_veg))/float(start_veg)
                percent_change=round(change_value*100,2)
            else:
                return "The vegetation index on both "+ str(start_date)+" and "+ str(end_date) + " is **"+ str(round(start_veg,3))+ "**. You can visualise this from the satellite images below "
            if change=="a decrease" or "an increase":
                message="The vegetation index on "+ str(start_date) +" is **" + str(round(start_veg,2)) +"** compared to **"+str((round(end_veg,2)))+ "** on "+ str(end_date) +\
                ". This represents **"+ str(change) + "** of **"+ str(percent_change)+"%**. For reference, the mean vegetation index before and after the pandemic is **"+str(prior_veg)+\
                "** and **"+str(during_veg)+"** respectively. You can visualise the difference from the satellite images below."  
                return message 
    except:
        return "Please ensure that the date inputs are not in the future. Otherwise, information on the requested date is not available"
    

def date_select2(date_input):    
    try:
        dt = datetime.datetime.strptime(date_input,'%Y-%m-%d').date()
        d=dt.strftime('%Y-%m-%d')
        year_no=int(d[0:4])-2018
        date_object = int(dt.timetuple().tm_yday)
        n= ((date_object-1)//16)
        with open(r'/home/shawn/Downloads/NDVI.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
        if dt<=datetime.date.today() and dt>=datetime.datetime.strptime('01/01/2018','%m/%d/%Y').date():
            if dt<=datetime.datetime.strptime(data[58][0],'%Y/%m/%d').date():
                return data[n+1+year_no*23][0]
            else:
                return data[58][0]
        elif dt>datetime.date.today() or dt<datetime.datetime.strptime('01/01/2018','%m/%d/%Y').date():
            return "null"
    except ValueError:
        return "null"

def image_output2(start_date, end_date):
    start_year=date_select2(start_date)[0:4]
    start_month=date_select2(start_date)[5:7]
    start_day=date_select2(start_date)[8:10]
    start_date2=start_year+'-'+start_month+'-'+start_day
    
    end_year=date_select2(end_date)[0:4]
    end_month=date_select2(end_date)[5:7]
    end_day=date_select2(end_date)[8:10]
    end_date2=end_year+'-'+end_month+'-'+end_day
    
    image_dir=r'/home/shawn/NDVI_Images'  
    start_img=image_dir+'/'+str(start_date2)+'.png'
    print(start_img)
    end_img=image_dir+'/'+str(end_date2)+'.png'
    print(end_img)
    req_img=[start_img,end_img]
    all_img=[]
    img_path=[]
    for file in glob.iglob(image_dir + '**/*'):
        all_img.append("".join(file))
    for img in req_img:
        if img in all_img:
            img_path.append(img)
    
    if len(img_path)<2 or len(NDVI_compare(start_date, end_date))<120:
        return "null"
    
    elif len(img_path)==2:
        #start image
        
        im = Image.open(img_path[0]) 
        draw = ImageDraw.Draw(im)
        font = ImageFont.truetype("arial.ttf", 25)
        (x, y) = (500, 900)
        message = str(start_date)
        color = 'rgb(0, 0, 0)' # black color
        # draw the message on the background
        draw.text((x, y), message, fill=color, font=font)
        cropped = im.crop((200,450, 650, 950))
        
        #separator image
        im1_5=Image.open(r"/home/shawn/NDVI_Images/split2.png")
        cropped1_5 = im.crop((953, 450, 964, 950))
        
        #end image
        im2 = Image.open(img_path[1])
        draw = ImageDraw.Draw(im2)
        font = ImageFont.truetype("arial.ttf", 25)
        (x, y) = (500, 900)
        message = str(end_date)
        color = 'rgb(0, 0, 0)' # black color
        # draw the message on the background
        draw.text((x, y), message, fill=color, font=font)
        cropped2 = im2.crop((200,450, 650, 950))
        
        #sidebar image
        im3= Image.open(r"/home/shawn/NDVI_Images/sidebar_NDVI.png")
        cropped3= im3.crop((0,0,155,500))
        
        x=append_images([cropped,cropped1_5,cropped2,cropped3], direction='horizontal')
        x.save("/home/shawn/shared/NDVI_comparison.png")

def date_convert(date):
    dt = datetime.datetime.strptime(date, '%Y-%m-%d')
    start = dt - datetime.timedelta(days=dt.weekday()+1)
    end = start + datetime.timedelta(days=6)
    return start.strftime('%m/%d/%Y')


def forecast_pol(date):    
    try:
        with open(r'/home/shawn/Downloads/Pollution_forecast_values.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
            for row in data:
                if row[0]==date_convert(date):
                    return "The forecasted pollution level on "+ date+" is **"+ str(round(float(row[2]),2))+ " µmol/m2**."
    except:
        return "Forecasted value on selected date is not available."
    
def upper_pol(date):    
    try:
        with open(r'/home/shawn/Downloads/Pollution_forecast_values.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
            for row in data:
                if row[0]==date_convert(date):
                    return "The upper pollution level limit on "+ date+" is **"+ str(round(float(row[3]),2))+ " µmol/m2**."
    except:
        return "Forecasted value on selected date is not available."

def lower_pol(date): 
    try:
        with open(r'/home/shawn/Downloads/Pollution_forecast_values.csv') as csvDataFile:
            data=list(csv.reader(csvDataFile))
            for row in data:
                if row[0]==date_convert(date):
                    return "The lower pollution level limit on "+ date+" is **"+ str(round(float(row[1]),2))+ " µmol/m2**."
    except:
        return "Forecasted value on selected date is not available."

def forecast_ndvi(date):    
    try:
        with open(r'/home/shawn/Downloads/NDVI_forecast_values.csv', encoding="utf8") as csvDataFile:
            data=list(csv.reader(csvDataFile))
            month=date[5:7]
            year=date[0:4]
            for row in data:
                if row[0][0:2]==month and row[0][6:10]==year:

                    return "The forecasted Vegetation Index on "+ date+" is **"+ str(round(float(row[1]),2))+ "**."
    except:
        return "Forecasted value on selected date is not available."
    
def lower_ndvi(date):    
    try:
        with open(r'/home/shawn/Downloads/NDVI_forecast_values.csv', encoding="utf8") as csvDataFile:
            data=list(csv.reader(csvDataFile))
            month=date[5:7]
            year=date[0:4]
            for row in data:
                if row[0][0:2]==month and row[0][6:10]==year:

                    return "The lower Vegetation Index limit on "+ date+" is **"+ str(round(float(row[2]),2))+ "**."
    except:
        return "Forecasted value on selected date is not available."
    
def upper_ndvi(date):    
    try:
        with open(r'/home/shawn/Downloads/NDVI_forecast_values.csv', encoding="utf8") as csvDataFile:
            data=list(csv.reader(csvDataFile))
            month=date[5:7]
            year=date[0:4]
            for row in data:
                if row[0][0:2]==month and row[0][6:10]==year:

                    return "The upper Vegetation Index limit on "+ date+" is **"+ str(round(float(row[3]),2))+ "**."
    except:
        return "Forecasted value on selected date is not available." 

def respond_to_button_press(webhook):
    """Respond to a button press on the card we posted"""

    # Some server side debugging
    room = api.rooms.get(webhook.data.roomId)
    attachment_action = api.attachment_actions.get(webhook.data.id)
    person = api.people.get(attachment_action.personId)
    message_id = attachment_action.messageId
    print(
        f"""
        NEW BUTTON PRESS IN ROOM '{room.title}'
        FROM '{person.displayName}'
        """
    )
    def corr_image():
        JapanData=pd.read_csv('/home/shawn/Downloads/JPdataset_prefecture.csv', header= 0, encoding= 'unicode_escape')
        lst=['Infections', 'Deaths', 'Population', 'GDP', 'Healthcare', 'Pollution', 'Forest ratio', 'Temperature', 'Humidity %']
        params_lst=[]
        for params in lst:
            if attachment_action.json_data['inputs'][params]=='true':
                params_lst.append(params)
        JapanRevData=pd.DataFrame(JapanData[params_lst])
        sb.set(font_scale=2)
        f, axes=plt.subplots(1, 1, figsize=(20, 20))
        ax=sb.heatmap(JapanRevData.corr(), vmin = -1, vmax = 1, linewidths = 1, annot = True, fmt = ".2f", annot_kws = {"size": 25}, cmap = "RdBu") 
        ax.get_figure().savefig('/home/shawn/shared/output.png')
        ax2=sb.pairplot(data=JapanRevData)
        ax2.savefig('/home/shawn/shared/output1.png')
        
        
    print(attachment_action.json_data)
    if attachment_action.json_data["inputs"]["formDemoAction"]=="SubmitA":
        start_date=attachment_action.json_data["inputs"]["start_date"]
        end_date=attachment_action.json_data["inputs"]["end_date"]
        if attachment_action.json_data["inputs"]["Pollution"]=="true" and attachment_action.json_data["inputs"]["Forest Loss"]=="true":
            image_output(start_date, end_date)
            image_output2(start_date, end_date)
            if type(image_output(start_date, end_date)) and type(image_output2(start_date, end_date)) !=str:
                content = putvals(compare_response, [
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/pol_comparison.png",
            },
            {
                "key": "text",
                "target": "__TEXT1__",
                "val": f"{pollutionJP_datecompare(start_date,end_date)}",
            },
            {
                "key": "url",
                "target": "__DATA2__",
                "val": f"{webhook_url}/xdoc/NDVI_comparison.png",
            },
            {
                "key": "text",
                "target": "__TEXT2__",
                "val": f"{NDVI_compare(start_date,end_date)}",
            },
            ])
                ret = api.messages.create(
            roomId=room.id,
            text="If you see this your client cannot render cards",
            attachments=[{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": content
                    }],
            )
                debug_object(ret)
            else:
                api.messages.create(
                room.id,
                parentId=message_id,
                markdown=f"{pollutionJP_datecompare(start_date,end_date)}"
            )      
        elif attachment_action.json_data["inputs"]["Pollution"]=="true" and attachment_action.json_data["inputs"]["Forest Loss"]=="false":
            image_output(start_date, end_date)
            if type(image_output(start_date, end_date))!=str:
                content = putvals(compare_response2, [
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/pol_comparison.png",
            },
            {
                "key": "text",
                "target": "__TEXT1__",
                "val": f"{pollutionJP_datecompare(start_date,end_date)}",
            },
            ])
                ret = api.messages.create(
            roomId=room.id,
            text="If you see this your client cannot render cards",
            attachments=[{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": content
                    }],
            )
                debug_object(ret)
            else:
                api.messages.create(
                room.id,
                parentId=message_id,
                markdown=f"{pollutionJP_datecompare(start_date,end_date)}"
            )      
        elif attachment_action.json_data["inputs"]["Pollution"]=="false" and attachment_action.json_data["inputs"]["Forest Loss"]=="true":
            image_output(start_date, end_date)
            if type(image_output2(start_date, end_date))!=str:
                content = putvals(compare_response3, [
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/NDVI_comparison.png",
            },
            {
                "key": "text",
                "target": "__TEXT1__",
                "val": f"{NDVI_compare(start_date,end_date)}",
            },
            ])
                ret = api.messages.create(
            roomId=room.id,
            text="If you see this your client cannot render cards",
            attachments=[{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": content
                    }],
            )
                debug_object(ret)
            else:
                api.messages.create(
                room.id,
                parentId=message_id,
                markdown=f"{NDVI_compare(start_date,end_date)}"
            )                                

  
            
    if attachment_action.json_data["inputs"]["formDemoAction"]=="SubmitB":
        corr_image()
        content = putvals(insights_response, [
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/output.png",
            },
            {
                "key": "text",
                "target": "__TEXT1__",
                "val":f"Depicted above is the correlation heatmap of your chosen parameters. The values shown reflects the correlation between the intersecting parameters. Values range between 1.0 to -1.0, which reflects the degree of effect the change in one parameter has on the other.                                                          \n"
                     f"\n• A value between `0 to 0.3 (-0.3)` represents a **low degree** of correlation\n"  
                     f"\n• A value between `0.3 (-0.3) to 0.7 (-0.7)` represents a **moderate degree** of positive (negative) correlation\n"
                     f"\n• A value between `0.7 (-0.7) to 1.0 (-1.0)` represents a **strong degree** of positive (negative) correlation"
            },
            {
                "key": "url",
                "target": "__DATA2__",
                "val": f"{webhook_url}/xdoc/output1.png",
            },
            {
                "key": "text",
                "target": "__TEXT2__",
                "val":f"Depicted above is the 2 dimensional pairplot between your chosen parameters. Each point represents a prefecture of Japan. The data used for each parameter per prefecture is calculated using a two year average gathered from national census or satellite data."
            },
            ])
        ret = api.messages.create(
            roomId=room.id,
            text="If you see this your client cannot render cards",
            attachments=[{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": content
                    }],
            )
        debug_object(ret)
        
    if attachment_action.json_data["inputs"]["formDemoAction"]=="SubmitC":
        forecast_date=attachment_action.json_data["inputs"]["forecast_date"]
        if datetime.datetime.strptime(forecast_date, '%Y-%m-%d')<datetime.datetime.now():
            api.messages.create(
                room.id,
                parentId=message_id,
                markdown=f"Forecasted value on selected date is not available."
            )      
        
        elif attachment_action.json_data["inputs"]["ParamChoiceVal"]=="Pollution":
            if datetime.datetime.strptime(forecast_date, '%Y-%m-%d')>datetime.datetime(2020,12,26):
                api.messages.create(
                    room.id,
                    parentId=message_id,
                    markdown=f"Forecasted value on selected date is not available."
            )      
            else:
                content = putvals(forecast_response, [
            {
                "key": "text",
                "target": "__TEXT1__",
                "val": f"{forecast_pol(forecast_date)}"
                                              
            },
            {
                "key": "text",
                "target": "__TEXT1.5__",
                "val": f"With a confidence interval of 95%: \n"
                       f"\n{upper_pol(forecast_date)} {lower_pol(forecast_date)}"
                                              
            },
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/pol_forecast.png",
            },
            {
                "key": "text",
                "target": "__TEXT2__",
                "val":f"Shown above is the full weekly forecast for NO2 Pollution levels in your selected location. The blue line represents the recorded past pollution levels by the Sentinel 5P satellite. The green line represents future forecasted values, whilst the grey shaded region is the 95% confidence interval."
            },
            {               
                "key": "text",
                "target": "__TEXT3__",
                "val":f"The SARIMAX(0,1,1)x(0,1,[1,2],12) model was used for predicting NO2 Pollution levels. The parameters values chosen for the model was derived from auto_arima to yield the lowest error possible. A per week basis forecast was used as pollution data is only available from July 2018 onwards."
            },                
            {    
                "key": "text",
                "target": "__TEXT4__",
                "val":f"The Akaike Information Critera (AIC) value for this model is 1012.110. AIC is a widely used measure to quantify the goodness of fit and the simplicity/parsimony of the model into a single statistic."
            },
            ])
        elif attachment_action.json_data["inputs"]["ParamChoiceVal"]=="NDVI":
            if datetime.datetime.strptime(forecast_date, '%Y-%m-%d')>datetime.datetime(2021,5,31):
                api.messages.create(
                    room.id,
                    parentId=message_id,
                    markdown=f"Forecasted value on selected date is not available."
            )  
            else:
                content = putvals(forecast_response, [
            {
                "key": "text",
                "target": "__TEXT1__",
                "val": f"{forecast_ndvi(forecast_date)}"
                                             
            },
            {
                "key": "text",
                "target": "__TEXT1.5__",
                "val": f"With a confidence interval of 95%: \n"
                       f"\n{upper_ndvi(forecast_date)} {lower_ndvi(forecast_date)}"
                       
            },
            {
                "key": "url",
                "target": "__DATA1__",
                "val": f"{webhook_url}/xdoc/NDVI_forecast.png",
            },
            {
                "key": "text",
                "target": "__TEXT2__",
                "val":f"Shown above is the full monthly forecast for the Vegetation Index in your selected location. The blue line represents the recorded past Vegetation Indices by the MODIS 006 satellite. The green line represents future forecasted values, whilst the grey shaded region is the 95% confidence interval."
            },
            {               
                "key": "text",
                "target": "__TEXT3__",
                "val":f"Facebook Prophet was used for predicting Vegetation Indices. The model was chosen as it allows for seasonality modelling as well as accounts for holiday effects. Additional regressors such as precipitation was also added to improve its accuracy. The fourier order of 10 was chosen for seasonality to minimise fitting error. A per month basis forecast was used to account for long-term seasonal effects."
            },                
            {    
                "key": "text",
                "target": "__TEXT4__",
                "val":f"The Mean Absolute Percentage Error for this model is 20.118%. This measures the mean percentage error between all predicted values against its actual test data."
            },
            ])
        
        ret = api.messages.create(
            roomId=room.id,
            text="If you see this your client cannot render cards",
            attachments=[{
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": content
                    }],
            )
        debug_object(ret)    
        

def respond_to_message(webhook):
    """Respond to a message to our bot"""

    # Some server side debugging
    room = api.rooms.get(webhook.data.roomId)
    message = api.messages.get(webhook.data.id)
    person = api.people.get(message.personId)
    print(
        f"""
        NEW MESSAGE IN ROOM '{room.title}'
        FROM '{person.displayName}'
        MESSAGE '{message.text}'
        """
    )

    # This is a VERY IMPORTANT loop prevention control step.
    # If you respond to all messages...  You will respond to the messages
    # that the bot posts and thereby create a loop condition.
    if message.personId == me.id:
        # Message was sent by me (bot); do not respond.
        return "OK"

    else:
        # Message was sent by someone else; parse message and respond.
        if "/help" in message.text:
            api.messages.create(
                   room.id,
                   markdown=f"**Hi there! I'm a COVID-19 information bot. Here are my available functions:**\n"    	                 
                            f"\n`@Bot/help:` Details of how to use the bot\n"
                            f"\n`@Bot/features:` Details on what functionalities the bot has\n"
                            f"\n`@Bot/compare:` Activates the date comparison feature. The bot will send a card to collect your interested dates and parameters.\n" 
                            f"\n`@Bot/insights:` Activates the pandemic insights features. The bot will send a card to collect your interested parameters.\n"
                            f"\n`@Bot/forecast:` Activates the parameter forecast features. The bot will send a card to collect your interested date and parameter.\n"
                            f"\n **Note:** If you're using a 1:1 space with the bot, please exclude `@Bot` when calling functions." 
                   )
        if "/features" in message.text:
            api.messages.create(
                   room.id,
                   markdown=f"This bot is equipped with 3 functions. We have a **Date Comparison**, a **Pandemic Insights** feature, and a **Parameter Forecast** feature.\n"
                   f"\n**Date Comparison:** This feature allows you to enter different parameters of your liking, as well as a start and end date. You will then see a comparison between the selected parameters on those dates in both values and satellite images.\n"
                   f"\n**Pandemic Insights:** This feature allows you to enter different parameters of your liking. You will then be able to see the correlation of your chosen parameters with the COVID-19 situation in your chosen country.\n"
                   f"\n**Parameter Forecast:** This feature allows you to enter a parameter of your liking, as well as a date. You will then be able to see the forecasted value of your chosen parameter on the selected date together with a 95% confidence interval for your chosen location.")
                   
        if "/compare" in message.text:
            api.messages.create(
                   room.id,                   
                   text="If you see this your client cannot render cards",
                   attachments=[{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": CARD_DATECOMPARE
            }],
        )
        if "/insights" in message.text:
            api.messages.create(
                   room.id,                   
                   text="If you see this your client cannot render cards",
                   attachments=[{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": CARD_INSIGHTS
            }],
        )
        if "/forecast" in message.text:
            api.messages.create(
                   room.id,                   
                   text="If you see this your client cannot render cards",
                   attachments=[{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": CARD_FORECAST
            }],
        )

        return "OK"
        

# Core bot functionality
# Webex will post to this server when a message is created for the bot
# or when a user clicks on an Action.Submit button in a card posted by this bot
# Your Webex Teams webhook should point to http://<serverip>:<port>/events
@flask_app.route("/events", methods=["POST"])
def webex_teams_webhook_events():
    """Respond to inbound webhook JSON HTTP POST from Webex Teams."""
    # Create a Webhook object from the JSON data
    webhook_obj = Webhook(request.json)

    # Handle a new message event
    if (webhook_obj.resource == MESSAGE_WEBHOOK_RESOURCE
            and webhook_obj.event == MESSAGE_WEBHOOK_EVENT):
        respond_to_message(webhook_obj)

    # Handle an Action.Submit button press event
    elif (webhook_obj.resource == CARDS_WEBHOOK_RESOURCE
          and webhook_obj.event == CARDS_WEBHOOK_EVENT):
        respond_to_button_press(webhook_obj)
    
    elif (webhook_obj.resource == MEMBERSHIP_WEBHOOK_RESOURCE
          and webhook_obj.event == MEMBERSHIP_WEBHOOK_EVENT):
        respond_to_membership(webhook_obj)

    # Ignore anything else (which should never happen
    else:
        print(f"IGNORING UNEXPECTED WEBHOOK:\n{webhook_obj}")

    return "OK"

@flask_app.route("/xdoc/<path:path>", methods=["GET"])
def webex_teams_providing_documents(path):
    mimetypes = {
        ".png": "image/png",
        ".gif": "image/gif",
    }
    mimetype = mimetypes.get(os.path.splitext(path)[1])
    if mimetype is None:
        return "ERROR", 404
    try:
        content = open(os.path.join(shared_dir, path), "rb").read()
    except IOError as exc:
        return "ERROR", 404
    return Response(content, mimetype=mimetype)

def main():
    # Delete preexisting webhooks created by this script
    delete_webhooks_with_name()

    create_webhooks(webhook_url)

    try:
        # Start the Flask web server
        flask_app.run(host="0.0.0.0", port=port)

    finally:
        print("Cleaning up webhooks...")
        delete_webhooks_with_name()


if __name__ == "__main__":
    main()


