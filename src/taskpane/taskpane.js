/* global Office */

Office.onReady(() => {

    initialize();

});

function initialize() {

    // Confidence Slider

    const slider =
        document.getElementById(
            "confidenceSlider"
        );

    if (slider) {

        slider.addEventListener(
            "input",
            () => {

                const confidenceValue =
                    document.getElementById(
                        "confidenceValue"
                    );

                if (confidenceValue) {

                    confidenceValue.innerText =
                        slider.value + "%";
                }
            }
        );
    }

    // Auto Pilot Toggle

    const autoPilot =
        document.getElementById(
            "autoPilotToggle"
        );

    if (autoPilot) {

        autoPilot.addEventListener(
            "change",
            () => {

                const panel =
                    document.getElementById(
                        "autoPilotSettings"
                    );

                if (!panel) return;

                if (autoPilot.checked) {

                    panel.classList.remove(
                        "hidden"
                    );

                } else {

                    panel.classList.add(
                        "hidden"
                    );
                }
            }
        );
    }

    // Save Button

    const saveBtn =
        document.getElementById(
            "saveBtn"
        );

    if (saveBtn) {

        saveBtn.addEventListener(
            "click",
            saveSettings
        );
    }

    // Load Existing Settings

    loadSettings();
}

function saveSettings() {

    const tone =
        document.getElementById(
            "tone"
        );

    const autoPilot =
        document.getElementById(
            "autoPilotToggle"
        );

    const confidence =
        document.getElementById(
            "confidenceSlider"
        );

    const highConfidence =
        document.getElementById(
            "highConfidenceAction"
        );

    const lowConfidence =
        document.getElementById(
            "lowConfidenceAction"
        );

    const settings = {

        tone:
            tone
                ? tone.value
                : "Professional",

        autoPilot:
            autoPilot
                ? autoPilot.checked
                : false,

        confidence:
            confidence
                ? confidence.value
                : 80,

        highConfidenceAction:
            highConfidence
                ? highConfidence.value
                : "Save Draft",

        lowConfidenceAction:
            lowConfidence
                ? lowConfidence.value
                : "Notify User"
    };

    localStorage.setItem(
        "ai_reply_settings",
        JSON.stringify(settings)
    );

    showStatus(
        "Settings saved successfully."
    );
}

function loadSettings() {

    const saved =
        localStorage.getItem(
            "ai_reply_settings"
        );

    if (!saved) return;

    const settings =
        JSON.parse(saved);

    // Tone

    const tone =
        document.getElementById(
            "tone"
        );

    if (tone) {

        tone.value =
            settings.tone ||
            "Professional";
    }

    // Auto Pilot

    const autoPilot =
        document.getElementById(
            "autoPilotToggle"
        );

    if (autoPilot) {

        autoPilot.checked =
            settings.autoPilot || false;
    }

    // Confidence

    const confidence =
        document.getElementById(
            "confidenceSlider"
        );

    if (confidence) {

        confidence.value =
            settings.confidence || 80;
    }

    const confidenceValue =
        document.getElementById(
            "confidenceValue"
        );

    if (confidenceValue) {

        confidenceValue.innerText =
            (settings.confidence || 80)
            + "%";
    }

    // High Confidence Action

    const highConfidence =
        document.getElementById(
            "highConfidenceAction"
        );

    if (highConfidence) {

        highConfidence.value =
            settings.highConfidenceAction ||
            "Save Draft";
    }

    // Low Confidence Action

    const lowConfidence =
        document.getElementById(
            "lowConfidenceAction"
        );

    if (lowConfidence) {

        lowConfidence.value =
            settings.lowConfidenceAction ||
            "Notify User";
    }

    // Auto Pilot Panel

    const panel =
        document.getElementById(
            "autoPilotSettings"
        );

    if (
        panel &&
        settings.autoPilot
    ) {

        panel.classList.remove(
            "hidden"
        );
    }
}

function showStatus(message) {

    const statusBox =
        document.getElementById(
            "statusBox"
        );

    if (statusBox) {

        statusBox.innerText =
            message;
    }
}