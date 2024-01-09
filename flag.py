#--------------------------------------------This also good --------------------------#
import cv2
from pyzbar.pyzbar import decode
import time

# Global variables
snapshot = None
snap_button_pressed = False

def on_mouse(event, x, y, flags, param):
    global snapshot, snap_button_pressed

    if event == cv2.EVENT_LBUTTONDOWN:
        # Check if the click is within the Snap button region
        if 500 <= x <= 600 and 20 <= y <= 70:
            snap_button_pressed = True
            print("Snap button pressed!")

def trace_qr_code():
    global snapshot, snap_button_pressed

    # Open the default camera (index 0)
    cap = cv2.VideoCapture(0)

    # Wait for 5 seconds before starting QR code scanning
    time.sleep(5)

    # Create the main window
    cv2.namedWindow("Traced QR Code")

    while True:
        # Read a frame from the webcam
        ret, frame = cap.read()

        # Convert the frame to grayscale
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Display the Snap button
        cv2.rectangle(frame, (500, 20), (600, 70), (255, 255, 255), -1)
        cv2.putText(frame, "Snap", (530, 55), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 0), 2, cv2.LINE_AA)

        # Check for mouse events
        cv2.setMouseCallback("Traced QR Code", on_mouse)

        # Capture a snapshot if the Snap button is pressed
        if snap_button_pressed:
            snapshot = frame.copy()
            print("Snapshot captured!")

            # Decode QR codes in the snapshot
            qr_codes = decode(cv2.cvtColor(snapshot, cv2.COLOR_BGR2GRAY))

            # Iterate through detected QR codes
            for qr_code in qr_codes:
                rect = qr_code.rect

                # Extract the coordinates of the bounding box
                x, y, w, h = rect.left, rect.top, rect.width, rect.height

                # Draw the bounding box
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

                # Extract and print QR code data
                qr_data = qr_code.data.decode('utf-8')
                print(f"QR Code Data: {qr_data}")

                # Close the window after decoding the QR code
                cv2.destroyAllWindows()
                cap.release()
                return  # Exit the function immediately

            snap_button_pressed = False  # Reset the button state

        # Display the frame with traced QR codes
        cv2.imshow("Traced QR Code", frame)

        # Break the loop if 'q' key is pressed
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the webcam and close the window (in case 'q' is pressed)
    cap.release()
    cv2.destroyAllWindows()

# Call the function to start webcam QR code scanning
trace_qr_code()
