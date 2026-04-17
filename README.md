# Animal Approach-Avoidance Task (AAT) with Priming

This project implements an Animal Approach-Avoidance Task (AAT) designed to measure implicit motivational and affective responses toward animal stimuli under emotional priming conditions.

## Experimental Overview

Participants complete a behavioral task where they respond to images of animals following an emotional priming phase.

Each trial follows this sequence:

1. **Priming (text)**  
   A sentence related to **suffering** is displayed to induce an emotional context.

2. **Fixation cross (+)**  
   Brief attentional reset before stimulus presentation.

3. **Image presentation with cue**  
   An animal image appears with a visual cue (circle or square) at the center.

4. **Response (Approach / Avoid)**  
   Participants respond by:
   - **Approach (pull / down)** → simulates bringing the image closer  
   - **Avoid (push / up)** → simulates pushing the image away  

5. **Zoom feedback animation**  
   The image dynamically zooms in or out based on the response, reinforcing the motor action.

---

## Experimental Design

- **Total trials:** 220  
- **Blocks:** 2 (110 trials each)  
- **Priming:** Only *suffering-related* stimuli are used  
- **Stimuli:** 110 animal images (each presented twice)

Each image is shown under two conditions:
- Approach
- Avoid

### Block manipulation:
The mapping between cue and action is reversed between blocks:

- **Block 1:** Circle = Approach  
- **Block 2:** Circle = Avoid  

This prevents simple motor learning and forces cognitive processing of the cue.

---

## What the task measures

The task captures:

- Reaction times (RT)
- Accuracy of responses
- Approach vs avoidance tendencies

These measures are used to infer:
- Implicit affective biases
- Motivational responses to animal categories
- Modulation of behavior by emotional priming

---

## Physiological Integration

The experiment supports synchronization with physiological data using:

- **LSL (Lab Streaming Layer)** for real-time event markers
- Optional **serial triggers**

Markers are sent for:
- Priming onset
- Image onset (ID-based)
- Participant response
- Accuracy (correct/incorrect)

This allows integration with signals such as:
- EMG
- EEG
- Other biosensors

---

## Data Output

Each trial is logged into a CSV file including:

- Participant ID and condition (pre/post)
- Block and trial number
- Stimulus metadata (animal, group, ID)
- Priming content
- Cue type and expected response
- Actual response and accuracy
- Reaction time (ms)
- Input device (keyboard/joystick)
- Trigger values

---

## Input Methods

- Keyboard (↑ / ↓)
- Joystick (push / pull axis)

---

## Key Features of the Code

- Fully automated trial generation and balancing
- Deterministic randomization per participant
- Dynamic cue-response mapping per block
- Real-time trigger streaming via LSL
- Smooth visual feedback through animation
- Structured data logging for reproducibility

---

## Purpose

This implementation is designed for research in:

- Affective/cognitive neuroscience  
- Implicit cognition  
- Human–animal perception  
- Emotion-driven behavior  

It enables the study of how emotional context (suffering priming) influences automatic approach–avoidance behavior.
