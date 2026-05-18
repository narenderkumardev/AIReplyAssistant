/* global Office */

Office.onReady(() => {

    initialize();

});

function initialize() {

    initializeConfidenceSlider();

    initializeAutoReplyToggle();

    initializeSaveButton();

    loadSettings();
}

function initializeConfidenceSlider() {

    const slider =
        document.getElementById(
            "confidenceSlider"
        );

    if (!slider) return;

    slider.addEventListener(
        "input",
        () => {

            const confidenceValue =
                document.getElementById(
                    "confidenceValue"
                );

            if (!confidenceValue) return;

            confidenceValue.innerText =
                slider.value + "%";
        }
    );
}

function initializeAutoReplyToggle() {

    const autoReplyToggle =
        document.getElementById(
            "autoPilotToggle"
        );

    if (!autoReplyToggle) return;

    autoReplyToggle.addEventListener(
        "change",
        () => {

            toggleAutoReplyPanel(
                autoReplyToggle.checked
            );
        }
    );
}

function initializeSaveButton() {

    const saveBtn =
        document.getElementById(
            "saveBtn"
        );

    if (!saveBtn) return;

    saveBtn.addEventListener(
        "click",
        saveSettings
    );
}

function toggleAutoReplyPanel(isEnabled) {

    const panel =
        document.getElementById(
            "autoPilotSettings"
        );

    if (!panel) return;

    if (isEnabled) {

        panel.classList.remove(
            "hidden"
        );

    } else {

        panel.classList.add(
            "hidden"
        );
    }
}

function saveSettings() {

    try {

        const mailboxUser =
            Office.context.mailbox
            .userProfile.emailAddress;

        const settings = {

            mailboxUser:
                mailboxUser,

            tone:
                getElementValue(
                    "tone",
                    "Professional"
                ),

            autoReplyEnabled:
                getElementChecked(
                    "autoPilotToggle",
                    false
                ),

            confidence:
                getElementValue(
                    "confidenceSlider",
                    80
                ),

            businessHours:
                getElementValue(
                    "businessHours",
                    "Business Hours Only"
                ),

            endDate:
                getElementValue(
                    "endDate",
                    ""
                ),

            highConfidenceAction:
                getElementValue(
                    "highConfidenceAction",
                    "Save Draft"
                ),

            lowConfidenceAction:
                getElementValue(
                    "lowConfidenceAction",
                    "Notify User"
                ),

            knowledgeSources: {

                mailHistory:
                    getElementChecked(
                        "mailHistory",
                        true
                    ),

                sharepoint:
                    getElementChecked(
                        "sharepoint",
                        true
                    ),

                faqDb:
                    getElementChecked(
                        "faqDb",
                        true
                    )
            },

            cacheTime:
                new Date().toISOString()
        };

        localStorage.setItem(
            "ai_reply_settings",
            JSON.stringify(settings)
        );

        console.log(
            "Settings Saved:",
            settings
        );

        /*
        FUTURE REST API FLOW

        POST:
        /api/settings/save

        BODY:
        settings

        Backend:
        → Save to Supabase
        */

        showStatus(
            "Settings saved successfully."
        );

    }
    catch (error) {

        console.error(
            "Save Settings Error:",
            error
        );

        showStatus(
            "Failed to save settings."
        );
    }
}

function loadSettings() {

    try {

        const saved =
            localStorage.getItem(
                "ai_reply_settings"
            );

        if (!saved) return;

        const settings =
            JSON.parse(saved);

        setElementValue(
            "tone",
            settings.tone ||
            "Professional"
        );

        setElementChecked(
            "autoPilotToggle",
            settings.autoReplyEnabled ||
            false
        );

        setElementValue(
            "confidenceSlider",
            settings.confidence ||
            80
        );

        setElementText(
            "confidenceValue",
            (settings.confidence || 80)
            + "%"
        );

        setElementValue(
            "businessHours",
            settings.businessHours ||
            "Business Hours Only"
        );

        setElementValue(
            "endDate",
            settings.endDate || ""
        );

        setElementValue(
            "highConfidenceAction",
            settings.highConfidenceAction ||
            "Save Draft"
        );

        setElementValue(
            "lowConfidenceAction",
            settings.lowConfidenceAction ||
            "Notify User"
        );

        if (
            settings.knowledgeSources
        ) {

            setElementChecked(
                "mailHistory",
                settings
                .knowledgeSources
                .mailHistory
            );

            setElementChecked(
                "sharepoint",
                settings
                .knowledgeSources
                .sharepoint
            );

            setElementChecked(
                "faqDb",
                settings
                .knowledgeSources
                .faqDb
            );
        }

        toggleAutoReplyPanel(
            settings.autoReplyEnabled
        );

    }
    catch (error) {

        console.error(
            "Load Settings Error:",
            error
        );
    }
}

function showStatus(message) {

    const statusBox =
        document.getElementById(
            "statusBox"
        );

    if (!statusBox) return;

    statusBox.innerText =
        message;
}

function getElementValue(
    id,
    defaultValue
) {

    const element =
        document.getElementById(id);

    if (!element) {

        return defaultValue;
    }

    return element.value;
}

function getElementChecked(
    id,
    defaultValue
) {

    const element =
        document.getElementById(id);

    if (!element) {

        return defaultValue;
    }

    return element.checked;
}

function setElementValue(
    id,
    value
) {

    const element =
        document.getElementById(id);

    if (!element) return;

    element.value = value;
}

function setElementChecked(
    id,
    value
) {

    const element =
        document.getElementById(id);

    if (!element) return;

    element.checked = value;
}

function setElementText(
    id,
    value
) {

    const element =
        document.getElementById(id);

    if (!element) return;

    element.innerText = value;
}