# PowerPoint Control Using Hand Gestures

This project allows you to control a PowerPoint presentation using hand gestures detected through a webcam. The gestures are recognized using MediaPipe, and the PowerPoint application is controlled using the `pywin32` library.

## Features

- **Swipe Right**: Move to the next slide.
- **Swipe Left**: Move to the previous slide.
- **Pinch Gesture**: Close the presentation.

## Requirements

- Python 3.x
- Webcam
- PowerPoint installed on the system
- The following Python libraries:
  - `ctypes`
  - `pywin32`
  - `opencv-python`
  - `mediapipe`
  - `numpy`
  - `scipy`
  - `pillow`
  - `tensorflow`

## Installation

### Prerequisites

1. Ensure Python 3.x is installed on your system. You can download it from the [official website](https://www.python.org/).
2. Make sure you have a webcam connected and working.

### Installation Steps

1. Clone the repository or download the project files.
2. Navigate to the project directory in your terminal or command prompt.
3. Create a virtual environment (optional but recommended):

    ```bash
    python -m venv venv
    ```

4. Activate the virtual environment:

    - **Windows**:

        ```bash
        venv\Scripts\activate
        ```

    - **macOS/Linux**:

        ```bash
        source venv/bin/activate
        ```

5. Install the required Python packages:

    ```bash
    pip install -r requirements.txt
    ```

## How to Use

1. Open a terminal or command prompt in the project directory.
2. Run the main script:

    ```bash
    python main.py
    ```

3. The PowerPoint presentation specified in the script will open.
4. Use the following gestures to control the presentation:
    - **Swipe Right**: Move your thumb to the right of your index finger to go to the next slide.
    - **Swipe Left**: Move your index finger to the right of your thumb to go to the previous slide.
    - **Pinch Gesture**: Touch the thumb and index finger together to close the presentation.

5. Press the `Esc` key to exit the application manually.

## Troubleshooting

- **Webcam Issues**: Ensure the webcam is properly connected and accessible by the script. Check your drivers if you encounter issues.
- **PowerPoint Issues**: Verify that the PowerPoint file path is correct and that PowerPoint is installed on your system.
- **Gesture Recognition**: If gestures are not being recognized correctly, adjust the thresholds or lighting conditions. You can also modify the `pinch_threshold` in the script for better accuracy.

## Advanced Configuration

- **Threshold Adjustment**: The pinch gesture detection threshold can be modified in the script to fine-tune gesture recognition.
- **Custom Gestures**: You can add additional gestures by extending the hand landmarks processing logic.

## License

This project is licensed under the MIT License.
