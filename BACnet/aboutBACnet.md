# BACnet

"Building Automation and Control Network"

Object-Oriented, Open Architecture intended to provide a standard interface for Building Automation and Control networks

Developed in 19a87 by American Society of Heating, Refrigerating and Air-Conditioning Engineers (ASHRAE)

Concept is to capture the environment, similar to the human senses:
```
Sight
Hearing
Taste
Smell
Touch / Pain
Balance
Acceleration
Temperature
Kinaesthetic Sense (Movement)
```

BACnet is an unconnected, peer network. Any device can talk to other devices, there is no ongoing data transfer and communication is unscheduled without time critical operations.

Node data is interpreted in a variety of ways such as:
*Who-Is*
*Who-Has*
*I-Have*

## Supported Physical and Link Layer protocols
### PTP (Point to Point)
Communication over Modem and Voice Phone Lines
Slow: 9.8kbit/s -> 56.0kbit/s
Supports V.32bis and V.42 and Direct Cable Connection with EIA-232
### MS/TP (Master Slave/Token Passing)
SHielded twisted pairs, low cost.
### ARCNET
Token bus standard. Can run up to 7.5 Mbit/s
### Ethernet
Primary media in use today.

## Object Orientation
BACnet data is represented as *Objects*, *Properties* and *Services*

All BACnet *Objects* have an identifier (Such as Al-1) which allows BACnet to identify it.  

There are 60 standard object types.

BACnet Objects feature some of the following types:
```
Analogue Input
Analogue Output
Analogue Value
Binary Input
Binary Output
Binary Value
Multi-State Input
Multi-State Output
Calendar
Event-Enrolment
File
Notification-Class
Group
Loop
Program
Schedule
Command
Device
```

BACnet *Properties* are the values of objects, their are 123 defined properties of objects. Every object must contain an *Object-identifier*, *Object-name* and an *Object-type*

There are required classes for each object and anything else is at the discretion of the device manufacturer. Properties can be set as Read-Only or Read-Write at their discretion.

*Services* instruct BACnet objects or make requests for information for those objects. Services are how one BACnet device gets information from another. There are 32 defined services.

Different devices support different objects and different service requests. These are defined by the devices PICS (Protocol Implementation Conformance Statement)

## BACnet libraries

### BACPypes
