#!/usr/bin/env python3
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import os


class DefaultConfig:
    """ Bot Configuration """

    PORT = 3978
    APP_ID = os.environ.get("MicrosoftAppId", "bd4a8cbe-4d70-4b91-8c88-44924e845309")
    APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "CkXVV3jRQ_I-DGtpU6sdVk6d.O1.g_54W6")
