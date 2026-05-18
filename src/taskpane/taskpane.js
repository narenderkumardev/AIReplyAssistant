/* global Office */

const supabaseClient =
    window.supabase.createClient(

        "https://vpszsnlevsrphplciitx.supabase.co",

        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZwc3pzbmxldnNycGhwbGNpaXR4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg4NDE5MTksImV4cCI6MjA5NDQxNzkxOX0.zHm_4_css3VEIDqWGAi6oUrsDI9PdE-FkY5mvKBAZfU"
    );

Office.onReady(() => {

    initialize();

});

function initialize() {

    initializeConfidenceSlider();

    initializeAutoReplyToggle();

    initializeSaveButton();

    loadSettings();

    startAutoRefresh();
}

function startAutoRefresh() {

    setInterval(() => {

        loadSettings();

    }, 5000);
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

            setElementText(
                "confidenceValue",
                slider.value + "%"
            );
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
        async () => {

            toggleAutoReplyPanel(
                autoReplyToggle.checked
            );

            if (
                autoReplyToggle.checked
            ) {

                await setAutomationPending();
            }
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

async function setAutomationPending() {

    showStatus(
        "⚡ Auto Reply setup in progress..."
    );
}

async function saveSettings() {

    try {

        let mailboxUser =
            "unknown@local";

        if (
            Office.context &&
            Office.context.mailbox &&
            Office.context.mailbox.userProfile
        ) {

            mailboxUser =
                Office.context.mailbox
                .userProfile.emailAddress;
        }

        const autoReplyEnabled =
            getElementChecked(
                "autoPilotToggle",
                false
            );

        const settings = {

            mailboxUser:
                mailboxUser,

            tone:
                getElementValue(
                    "tone",
                    "Professional"
                ),

            autoReplyEnabled:
                autoReplyEnabled,

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
                    ""
                ) || null,

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

            automationStatus:
                autoReplyEnabled
                    ? "PENDING_SETUP"
                    : "NOT_ENABLED",

            automationMessage:
                autoReplyEnabled
                    ? "Preparing automation..."
                    : "Auto Reply disabled."
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

                automation_status:
                    settings.automationStatus,

                automation_message:
                    settings.automationMessage,

                automation_updated_at:
                    new Date(),

                updated_at:
                    new Date()
            });

        if (error) {

            console.error(
                "Supabase Save Error:",
                error
            );

            showStatus(
                "❌ Supabase save failed."
            );

            return;
        }

        showStatus(
            autoReplyEnabled
                ? "⚡ Auto Reply setup started."
                : "✅ Settings saved successfully."
        );

        if (autoReplyEnabled) {

            await triggerBackendSetup(
                mailboxUser
            );
        }
    }
    catch (error) {

        console.error(
            "Save Settings Error:",
            error
        );

        showStatus(
            "❌ Failed to save settings."
        );
    }
}

async function triggerBackendSetup(
    mailboxUser
) {

    try {

        /*
            FUTURE IMPLEMENTATION

            await fetch(
                "https://your-api/api/automation/setup",
                {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({
                        mailboxUser
                    })
                }
            );
        */

        console.log(
            "Backend setup placeholder:",
            mailboxUser
        );
    }
    catch (error) {

        console.error(
            "Backend setup error:",
            error
        );
    }
}

async function loadSettings() {

    try {

        let mailboxUser =
            "unknown@local";

        if (
            Office.context &&
            Office.context.mailbox &&
            Office.context.mailbox.userProfile
        ) {

            mailboxUser =
                Office.context.mailbox
                .userProfile.emailAddress;
        }

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
            .maybeSingle();

        let settings = data;

        if (!settings) {

            const local =
                localStorage.getItem(
                    "ai_reply_settings"
                );

            if (!local) {

                showStatus(
                    "Using default settings."
                );

                return;
            }

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

        updateAutomationStatus(
            settings
        );
    }
    catch (error) {

        console.error(
            "Load Settings Error:",
            error
        );

        showStatus(
            "Using local settings."
        );
    }
}

function updateAutomationStatus(
    settings
) {

    const status =
        settings.automation_status ||
        settings.automationStatus;

    const message =
        settings.automation_message ||
        settings.automationMessage;

    if (!status) {

        showStatus(
            "Ready"
        );

        return;
    }

    switch (status) {

        case "PENDING_SETUP":

            showStatus(
                "⚡ " + message
            );

            break;

        case "INITIALIZING":

            showStatus(
                "⏳ " + message
            );

            break;

        case "ACTIVE":

            showStatus(
                "✅ Auto Reply Active"
            );

            break;

        case "ERROR":

            showStatus(
                "❌ " + message
            );

            break;

        default:

            showStatus(
                message || "Ready"
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