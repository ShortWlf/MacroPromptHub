MacroPromptHub
The "Context-Aware" Macro Execution Engine

MacroPromptHub is a lightweight C# utility that bridges the gap between static text snippets and automated workflow execution. It allows users to define custom buttons that not only store text but also "know" exactly which window to send that text to.

Key Features
Targeted Window Focus: The app doesn't just "paste." It uses a Window Picker to identify specific process names and window titles. It stores these in a Process|Title format to ensure your macros hit the exact browser tab or document you intended.

Physical Input Simulation: By utilizing SendKeys and native Windows API calls (SetForegroundWindow), the app simulates a user physically typing. This allows it to bypass limitations of simple copy-paste by triggering "Enter" commands and keyboard shortcuts within the target app.

Dynamic Hotkey System: Users can enter "Hotkey Assign Mode" to bind any physical button on their keyboard to a macro. These hotkeys are registered globally, allowing you to trigger actions even when the application is minimized to the system tray.

Multi-Mode Macros: * App Macros: Focuses a specific window, pastes text, and hits enter.

Browser Macros: Launches specific URLs and waits for the page to load before injecting prompts (perfect for AI tools like Copilot or ChatGPT).

Run Macros: Acts as a quick-launch bar for applications and folders.

Persistent Configuration: All settings and macros are saved in a human-readable prompts.txt file, making it easy to back up, share, or edit manually.

Technical Highlights
Language: C# / .NET WinForms

Automation: Simulates CTRL+V and {ENTER} via SendKeys.SendWait.

Native Interop: Uses user32.dll for "Nuclear" window focusing, ensuring the target app is pulled to the foreground before execution.

UI: A clean, dynamic FlowLayoutPanel that scales as you add more buttons.

How it Works
Select a Window: Use the built-in picker to find a running process (e.g., Chrome, Notepad, or a Terminal).

Define the Action: Input the text you want to send and whether you want the macro to "Press Enter" afterward.

Execute: Click the button or hit your assigned Hotkey. MacroPromptHub will find the window, bring it to the front, and "type" the data for you.
