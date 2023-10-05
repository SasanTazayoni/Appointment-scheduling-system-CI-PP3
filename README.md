# Appointment Scheduling System (Prototype)

## Introduction

This application is part of a future full-stack solution designed for a real IT support consultant. The consultant currently manages bookings through emails and phone calls, which leads to various issues, such as last-minute cancellations and difficulty in keeping track of booking times due to frequent changes.

The purpose of this application is to enable the client to efficiently handle and plan their bookings in advance. While the current version lacks detailed information for each booking, a future iteration envisions a robust full-stack solution that will revolutionise the way the IT support consultant manages their business operations.

In these upcoming iterations, the application aims to integrate a multitude of features to enhance both the consultant's workflow and the overall client experience. Some of the key enhancements include:

**Automated Reminders:** To reduce the occurrence of last-minute cancellations and no-shows, the system will implement automated appointment reminders. These reminders will be sent to the clients via email and text to mobiles, ensuring that they are well-informed about upcoming scheduled appointments.

**A comprehensive Database:** The future iteration will include a SQL database system to store and organise all booking details. This database will not only provide permanent access to booking records but also record service history as well as appointment details, allowing the consultant to track past services, issues resolved, and recommended follow-up actions.

**A website:** A user-friendly website will be developed to allow clients to register, log in, and book appointment requests online. Clients can conveniently browse the consultant's availability, select suitable time slots, and submit their appointment requests, improving accessibility and convenience for clients.

**A deposit System:** To mitigate last-minute cancellations and ensure commitment from clients, the application will introduce a deposit system. Clients may be required to make a deposit when booking appointments, which can be refunded or applied towards the service fee upon completion of the appointment.

**Excel spreadsheets:** Used instead of Google Sheets due to the ease of adding appointment notes. Adding notes to booked appointments is crucial, and Excel allows me to achieve this seamlessly through Python scripting.

These planned enhancements will transform the application into a comprehensive, end-to-end solution that not only addresses the consultant's current challenges but also takes their business to the next level of efficiency and customer service. By providing a seamless, integrated platform for appointment management, billing, communication, and client engagement, the consultant can focus more on delivering top-notch IT support while leaving the administrative tasks to the application.

You can access the application [here](https://appointment-booking-system-987e0f827702.herokuapp.com/). <br>
You can also access the associated spreadsheet [here](https://docs.google.com/spreadsheets/d/1uBX51j8qVqieYV65oMwpC26L3HZVyuoybPL4kNRlyxY/edit?usp=sharing).

***IMPORTANT NOTE: The login credentials are 'admin' for the username and 'password' for the password.***

## How to use the application

* The user is asked whether they would like to login (typing 'n' exits the application)
* A login consists of a username and password and these need to be typed correctly in order to log in (details provided above).
* After logging in, the program detects the time and date of log in, where each week starts on a Monday at 12:00 AM, and automatically updates the spreadsheet for week1 to match the current week and all the following weeks are shifted along accordingly. Previous weeks are deleted because they are no longer relevant and new weeks with open slots are appended to the end of the spreadsheet. If the user logs in within the current week, no update takes place. In a future iteration of this application, the deleted weeks will be archived in a database for reference.
* After this, the user is presented with a set of weeks from 1-12 with numbers allocated to each ("week1" represents the current week) and an option to exit the application.
* After selecting an appropriate week, the user is then presented with the days of that week with numbers allocated to each day. There is also an option to return to the week selection in case of a mistake.
* After selecting a day, the user is then presented the time slots for the selected day and whether the time slots are booked, blocked or open according to the spreadsheet. The user now has the option to either exit the program altogether, to cancel the choice which allows the user to reselect a day, to select a specific time slot or to select a range of time slots. Selecting a single time slot allows the alteration that single slot only, whereas selecting multiple slots allows multiple slots to be changed.
* If the user selects an open slot, or multiple open slots, they have the choice to book these slots, block these slots or to cancel their choice and return to the previous menu to select different slots. Booking a time slot or range, automatically puts a block at the start and end of the bookings so that appointments are not booked back to back and the consultant has time to travel to the appointment destination.
* If the user selects a booked slot or multiple booked slots, they have the option of either cancelling them all or returning to the previous menu.
* If the user selects a blocked slot or multiple blocked slots, they have the option of either unblocking them all or returning to the previous menu.
* If the user selects a mixture of open and blocked slots, they have the option of either blocking all the open slots, unblocking the blocked slots or returning to the previous menu.
* If the user selects a mixture of open and booked slots, they have the option of either booking all the open slots, cancelling the booked slots or returning to the previous menu.
* If the user selects a mixture of blocked and booked slots, since booked appointments cannot be blocked and blocked slots cannot be booked the user has only 2 options: cancel all appointments and unblock all slots or return to the previous menu.
* If the user selects a mixture of blocked, booked and open slots, the same logic as above applies where the user has only 2 options: cancel all appointments and unblock all slots or return to the previous menu.
* For each change of slot, there is a confirmation which if not confirmed, does not change the slot and reprompts the user for the appropriate action.
* If the user confirms, the program then adjusts the spreadsheet accordingly.
* When a slot has been changed, the user is prompted as to whether they would like to schedule more appointments. If yes, the program loops to the day select again. If no, the program exits.

**Please note: Although these instructions might appear complex, the application is actually straightforward and user-friendly.**

## User stories

As a user I want to:

* Have a secure system which only I can access so that I can manage my weekly work schedule.
* View my current schedule for any specific day.
* Select a week, day and time to book an appointment.
* Ensure that my bookings are not booked back-to-back so that I have time to plan my journey and travel to my next customer.
* Block slots for times on days where I am busy with other endeavours and unblock slots when necessary.
* Cancel appointments that are no longer required.
* Manage multiple slots to book long sessions, block entire days or cancel multiple appointments and slots.
* Navigate the system with ease and have feedback for every decision that I make.
* Cancel my decision if I make any mistakes at any point while I am using the application.
