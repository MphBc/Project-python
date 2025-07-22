############################################
### python file name: notify_ms_teams.py ###
############################################
import json
import requests
import os
import logging as log

HTTP_STATUS_OK = [200, 201, 202, 204]

def create_adaptive_card(header_text="MS Teams Notification", message_body="Sample message body"):
    """Create and return an adaptive card."""
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "text": header_text,
                "style": "heading",
                "size": "Large",
                "weight": "bolder",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "weight": "default",
                "wrap": True,
                "size": "default",
                "text": message_body,
            },
        ],
    }


def send_adaptive_card_to_ms_teams(webhook, adaptive_card):
    """Send an adaptive card to an MS Teams channel using a webhook."""
    payload = json.dumps(
        {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": adaptive_card,
                }
            ],
        }
    )
    headers = {"Content-Type": "application/json"}
    response = requests.post(webhook, data=payload, headers=headers)


    if response.status_code in HTTP_STATUS_OK:
        log.info(f"[{response.status_code}] Successfully called send notification to MS Team")
    else:
        log.error(f"[{response.status_code}] Failed to send message to MS Team Webhook API")
        log.error(
            f"Status code: {response.status_code}, Response: {response.text}"
        )



# Example usage
if __name__ == "__main__":
    # Load environment variables from .env file
    from dotenv import load_dotenv
    load_dotenv(override=True)

    webhook_url = os.getenv("MS_TEAMS_WEBHOOK", "YOU_MS_TEAMS_WEBHOOK_HERE")
    print(f"Webhook URL: {webhook_url}")
    body = os.getenv("SAMPLE_MESSAGE_BODY", "Sample message body")
    adaptive_card = create_adaptive_card(message_body=body)
    send_adaptive_card_to_ms_teams(webhook_url, adaptive_card)
