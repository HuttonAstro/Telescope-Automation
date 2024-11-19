"""All Conversion functions going between the angle in degrees and arc time"""
import math

def degreetohour(angle):
    """ Used to convert angles in degrees into arc time (hours).
    The equation to convert into hour angles from degrees is: Arc Time (Hours) = Angle (Degrees) / 15 """
    hourangle = angle / 15
    return hourangle

def degreetominute(angle):
    """ Used to convert angles in degrees into arc time (minutes).
    The equation to convert into hour angles from degrees is: Arc Time (minutes) = Angle (Degrees) * 60. """
    minuteangle = angle * 60
    return minuteangle

def degreetomseconds(angle):
    """ Used to convert angles in degrees into arc time (hours).
     The equation to convert into hour angles from degrees is: Arc Time (seconds) = Angle (Degrees) * 3600. """
    secondangle = angle * 3600
    return secondangle

def hourtodegree(arctime_hours):
    """" Used to convert angles from arc time (hours) to degrees.
    The equation to convert into degrees from arc time is: Angle (Degrees) = Arc Time (Hours) * 15. """
    degrees_hour = arctime_hours * 15
    return degrees_hour

def minutetodegree(arctime_minutes):
    """ Used to convert angles from arc time (minutes) to degrees.
    # The equation to convert into degrees from arc time is: Angle (Degrees) = Arc Time (minutes) / 60. """
    degrees_minute = arctime_minutes / 60
    return degrees_minute

def secondtodegree(arctime_seconds):
    """Used to convert angles from arc time (seconds) to degrees.
    The equation to convert into degrees from arc time is: Angle (Degrees) = Arc Time (seconds) / 3600. """
    degrees_seconds = arctime_seconds / 3600
    return degrees_seconds