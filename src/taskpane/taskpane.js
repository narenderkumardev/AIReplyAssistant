/* global Office */

Office.onReady(() => {

    initialize();

});

function initialize() {

    const slider =
        document.getElementById(
            "confidenceSlider"
        );

    slider.addEventListener(
        "input",
        () => {

            document.getElementById(
                "confidenceValue"
            ).innerText =
                slider.value + "%";
        }
    );

    const autoPilot =
        document.getElementById(
            "autoPilotToggle"
        );

    autoPilot.addEventListener(
        "change",
        () => {

            const panel =
                document.getElementById(
                    "autoPilotSettings"
                );

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

    document.getElementById(
        "saveBtn"
    ).addEventListener(
        "click",
        saveSettings
    );

    loadSettings();
}

function saveSettings() {

    const settings = {

        tone:
            document.getElementById(
                "tone"
            ).value,

        autoPilot:
            document.getElementById(
                "autoPilotToggle"
            ).checked,

        confidence:
            document.getElementById(
                "confidenceSlider"
            ).value,

        highConfidenceAction:
            document.getElementById(
                "highConfidenceAction"
            ).value,

        lowConfidenceAction:
            document.getElementById(
                "lowConfidenceAction"
            ).value
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

    document.getElementById(
        "tone"
    ).value =
        settings.tone ||
        "Professional";

    document.getElementById(
        "autoPilotToggle"
    ).checked =
        settings.autoPilot || false;

    document.getElementById(
        "confidenceSlider"
    ).value =
        settings.confidence || 80;

    document.getElementById(
        "confidenceValue"
    ).innerText =
        (settings.confidence || 80)
        + "%";

    document.getElementById(
        "highConfidenceAction"
    ).value =
        settings.highConfidenceAction ||
        "Save Draft";

    document.getElementById(
        "lowConfidenceAction"
    ).value =
        settings.lowConfidenceAction ||
        "Notify User";

    if (settings.autoPilot) {

        document.getElementById(
            "autoPilotSettings"
        ).classList.remove(
            "hidden"
        );
    }
}

function showStatus(message) {

    document.getElementById(
        "statusBox"
    ).innerText =
        message;
}