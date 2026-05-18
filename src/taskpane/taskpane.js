

const supabaseClient =
    window.supabase.createClient(

        process.env.SUPABASE_URL,

        process.env.SUPABASE_ANON_KEY
    );

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

async function saveSettings() {

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
                parseInt(
                    getElementValue(
                        "confidenceSlider",
                        80
                    )
                ),

            businessHours:
                getElementValue(
                    "businessHours",
                    "Business Hours Only"
                ),

            endDate:
                getElementValue(
                    "endDate",
                    null
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
            }
        };

        localStorage.setItem(
            "ai_reply_settings",
            JSON.stringify(settings)
        );

        const { error } =
            await supabaseClient
            .from("user_settings")
            .upsert({

                mailbox_user:
                    settings.mailboxUser,

                tone:
                    settings.tone,

                auto_reply_enabled:
                    settings.autoReplyEnabled,

                confidence:
                    settings.confidence,

                business_hours:
                    settings.businessHours,

                end_date:
                    settings.endDate,

                high_confidence_action:
                    settings.highConfidenceAction,

                low_confidence_action:
                    settings.lowConfidenceAction,

                knowledge_sources:
                    settings.knowledgeSources,

                updated_at:
                    new Date()
            });

        if (error) {

            console.error(
                "Supabase Save Error:",
                error
            );

            showStatus(
                "Supabase save failed."
            );

            return;
        }

        showStatus(
            "Settings synced successfully."
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

async function loadSettings() {

    try {

        const mailboxUser =
            Office.context.mailbox
            .userProfile.emailAddress;

        const {
            data,
            error
        } = await supabaseClient
            .from("user_settings")
            .select("*")
            .eq(
                "mailbox_user",
                mailboxUser
            )
            .single();

        let settings = data;

        if (!settings) {

            const local =
                localStorage.getItem(
                    "ai_reply_settings"
                );

            if (!local) return;

            settings =
                JSON.parse(local);
        }

        if (error) {

            console.log(
                "Using local cache settings."
            );
        }

        setElementValue(
            "tone",
            settings.tone ||
            "Professional"
        );

        setElementChecked(
            "autoPilotToggle",
            settings.auto_reply_enabled ||
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
            settings.business_hours ||
            settings.businessHours ||
            "Business Hours Only"
        );

        setElementValue(
            "endDate",
            settings.end_date ||
            settings.endDate ||
            ""
        );

        setElementValue(
            "highConfidenceAction",
            settings.high_confidence_action ||
            settings.highConfidenceAction ||
            "Save Draft"
        );

        setElementValue(
            "lowConfidenceAction",
            settings.low_confidence_action ||
            settings.lowConfidenceAction ||
            "Notify User"
        );

        toggleAutoReplyPanel(
            settings.auto_reply_enabled ||
            settings.autoReplyEnabled
        );

        showStatus(
            "Settings loaded successfully."
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