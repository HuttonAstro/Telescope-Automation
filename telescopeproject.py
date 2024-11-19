from win32com.client import Dispatch
import json
import time
import Conversions
from tkinter import *
from datetime import datetime
from astropy import units as u
from astropy.time import Time
from astropy.coordinates import SkyCoord, FK5
import math
import os.path

camstatus = {0: 'Camera is Not Connected.', 1: 'Camera is Reporting an Error.', 2: 'Connected but Idle.',
                     3: 'Camera is exposing a light image.', 4: 'Camera is reading an image from the sensor array.',
                     5: 'Camera is downloading an image to the computer.',
                     6: 'Camera is flushing the sensor array', 7: 'Camera is waiting for a trigger signal.',
                     8: 'Camera Control Window is waiting for MaxIm DL to be ready to accept an image.',
                     9: 'Camera Control is waiting for it to be time to acquire next image.',
                     10: 'Camera is exposing a dark needed by Simple Auto Dark',
                     11: 'Camera is exposing a Bias frame', 12: 'Camera is exposing a Dark frame',
                     13: 'Camera is exposing a Flat frame',
                     15: 'Camera Control window is waiting for filter wheel or focuser'}
filters = {'Blue': 1, 'Moon': 2, 'Green': 3, 'Red': 4, 'Yellow': 5, 'Clear': 6, 'None': 0}
filters_c = {1: 'Blue', 2: 'Moon', 3: 'Green', 4: 'Red', 5: 'Yellow', 6: 'Clear', 0: 'None'}


class Automation:

    def __init__(self, CameraName, TelescopeName, textfile_name, temp, Image_save_path):
        self.cam = None
        self.tel = None
        self.save = Image_save_path + "\\" + str(
            datetime.now().strftime("%Y-%m-%d"))  # Establish a savepath with the current date.
        self.set_temp = temp
        self.text = textfile_name
        if self.tel is None:
            self.ConnectToTelescope(
                TelescopeName)  # Call the ConnectToTelescope function to establish a connection to the telescope.
        if self.cam is None:
            self.ConnectToCamera(CameraName,
                                 self.set_temp)  # Call the ConnectToCamera function to establish a connection to the camera and set the desired temperature.
        if not os.path.exists(Image_save_path):
            os.mkdir(Image_save_path)  # If the parent folder does not exist, it creates the folder.
        if not os.path.exists(self.save):
            os.mkdir(self.save)  # If the savepath does not exist, it creates the savepath within the parent folder.
        with open(self.text, 'r') as infile:  # Open the json file and read the contents.
            self.jsonDic = json.load(infile)  # Localizing the dictionary from the json file.
        self.target_list = self.sorted_targets()  # Creates a sorted target list by calling the sortedtargets function.

        # Building the GUI
        self.mp = Tk()
        self.mp.title('Telescope Automation')
        mainlabel = Label(self.mp, text='Telescope Vitals')
        mainlabel.config(fg='black', bg="steelblue")
        mainlabel.config(font=('Helvetica', 20, 'bold'))
        mainlabel.config(height=1, width=40)
        mainlabel.grid(row=0, column=0, columnspan=10, sticky=N + S + W + E)

        # UTC Label
        labeltime = Label(self.mp)
        labeltime.grid(row=1, column=8, sticky=E + W)
        labeltime.config(text='Time (UTC):', font=('Helvetica', 12), bg='SteelBLue1')
        self.phase = Label(self.mp)
        self.phase.grid(row=1, column=9, sticky=W + E)
        dt = datetime.utcnow().strftime("%H:%M:%S")
        self.phase.config(text=dt, font=("Helvetica", 12), bg='lightslategrey')

        # RA Labels
        labelRA = Label(self.mp)
        labelRA.grid(row=1, column=0, sticky=E + W)
        labelRA.config(text='Current RA:', font=('Helvetica', 12), bg='SteelBLue1')
        self.right = Label(self.mp, width=6)
        self.right.grid(row=1, column=1, sticky=W + E)
        self.RA = self.tel.RightAscension
        self.RA = round(self.RA, 4)
        self.right.config(text=self.RA, font=("Helvetica", 12), bg='slategrey')

        # Declination Labels
        labeldec = Label(self.mp)
        labeldec.grid(row=1, column=2, sticky=E + W)
        labeldec.config(text='Current Dec:', font=('Helvetica', 12), bg='SteelBLue1')
        self.Decline = Label(self.mp, width=7)
        self.Dec = self.tel.Declination
        self.Dec = round(self.Dec, 3)
        self.Decline.grid(row=1, column=3, sticky=E + W)
        self.Decline.config(text=self.Dec, font=("Helvetica", 12), bg='slategrey')

        #Filter Wheel Labels
        filterlabel = Label(self.mp)
        filterlabel.grid(row=2, column=6, sticky=E+W)
        filterlabel.config(text="Current Filter:", font=("Helvetica",12),bg='SteelBLue1')
        self.fil = Label(self.mp, width=10)
        self.fil.grid(row=2,column=7,sticky=W+E)
        self.color = self.cam.Filter
        self.fil.config(text=filters_c[self.color], font=("Helvetica", 12),bg='slategrey')

        # CCD Labels
        camlabel = Label(self.mp)
        camlabel.grid(row=3, column=0, sticky=E + W)
        camlabel.config(text="Camera Status:", font=("Helvetica", 12), bg='SteelBLue1')
        self.camcom = Label(self.mp, width=80)
        self.camcom.grid(row=3, column=1, columnspan=9, sticky=W + E)
        self.camstat = camstatus[self.cam.CameraStatus]
        self.camcom.config(text=self.camstat, font=("Helvetica", 12), bg='Green2')

        # CCD Camera Temperature Labels
        templabel = Label(self.mp)
        templabel.grid(row=1, column=6, sticky=E + W)
        templabel.config(text="Temperature (C):", font=("Helvetica", 12), bg='SteelBLue1')
        self.temp = Label(self.mp)
        self.temp.grid(row=1, column=7, sticky=E + W)
        self.temp.config(text=round(self.cam.Temperature, 1), font=("Helvetica", 12), bg='IndianRed1')

        # Telescope Status Labels
        scopelabel = Label(self.mp)
        scopelabel.grid(row=1, column=4, sticky=E + W)
        scopelabel.config(text="Telescope Status:", font=("Helvetica", 12), bg='SteelBLue1')
        self.scopestat = Label(self.mp, width=15)
        self.scopestat.grid(row=1, column=5, sticky=W + E)
        self.scopestat.config(text="Not Slewing", font=('Helvetica', 12), bg='green2')

        # Target Labels
        targetlabel = Label(self.mp)
        targetlabel.grid(row=2, column=0, sticky=E + W)
        targetlabel.config(text="Current Target:", font=("Helvetica", 12), bg='SteelBlue1')
        self.target = Label(self.mp, width=10)
        self.target.grid(row=2, column=1, sticky=W + E)
        self.target.config(text="None", font=("Helvetica", 12), bg='slategrey')

        # Next Target Labels
        nextlabel = Label(self.mp)
        nextlabel.grid(row=2, column=2, sticky=E + W)
        nextlabel.config(text="Next Target:", font=("Helvetica", 12), bg='SteelBlue1')
        self.next = Label(self.mp, width=10)
        self.next.grid(row=2, column=3, sticky=W + E)
        self.next.config(text="None", font=("Helvetica", 12), bg='slategrey')

        # Number of Targets Labels
        numtarlabel = Label(self.mp)
        numtarlabel.grid(row=2, column=4, sticky=E + W)
        numtarlabel.config(text="Targets:", font=("Helvetica", 12), bg='SteelBlue1')
        self.target_number = Label(self.mp)
        self.target_number.grid(row=2, column=5, sticky=E + W)
        self.target_number.config(text="None", font=("Helvetica", 12), bg='slategrey')

        # Number of Exposures Labels
        numberlabel = Label(self.mp)
        numberlabel.grid(row=2, column=8, sticky=E + W)
        numberlabel.config(text="Exposures:", font=("Helvetica", 12), bg='SteelBlue1')
        self.light = Label(self.mp)
        self.light.grid(row=2, column=9, sticky=E + W)
        self.light.config(text="None", font=("Helvetica", 12), bg='slategrey')

        self.var = StringVar(self.mp)
        self.var.set("Pause")
        self.w = OptionMenu(self.mp, self.var, "Run", "Pause")
        self.w.grid(row=4, column=3)
        self.k = 0
        self.u = 0

        self.refresh()  # Calls the refresh function.
        mainloop()

    def refresh(self):
        self.value_update()
        self.temp_update()
        self.status_update()
        self.phase.after(1000, self.refresh)
        if self.var.get() == "Run":
            self.running()
        elif self.var.get() == "Pause":
            self.w.config(bg='firebrick')  # Sets pause/play button color to red.
            # Arms restart button.
            Button(self.mp, text='Start Over', command=self.restart, bg="red").grid(row=4, column=6)
            self.tel.AbortSlew  # Stops from slewing if it is currently slewing.
            self.cam.AbortExposure()  # Stops the camera from taking an exposure.
            self.u = 0

    def sorted_targets(self):
        """
        Sorts the Target list to make slewing more efficient.
        """
        target_list = []
        d_o = 360
        target = None
        for key in self.jsonDic.keys():  # Looking for the target with the smallest right ascension.
            d = Conversions.hourtodegree(
                self.jsonDic[key]["RA"])  # Convert the right ascension from arc hours to degrees.
            if d < d_o:  # If the current right ascension is less then our current smallest value then we will make it our newest smallest value.
                d_o = d  # Sets our new smallest RA to be d_o.
                target = key  # Setting the target which is the key in the dictionary to be our target value.
        target_list.append(
            target)  # Appends our inital target value into our sorted target list once going through all of them and finding this has the lowest RA.
        for i in range(len(self.jsonDic.keys()) - 1):
            s_o = None
            next_target = None
            for key in self.jsonDic.keys():
                if key not in target_list:  # Checking to see if the value is already in our list or not.
                    d_ra = abs(Conversions.hourtodegree(int(self.jsonDic[target_list[-1]]["RA"]))
                               - Conversions.hourtodegree(
                        int(self.jsonDic[key]["RA"])))  # Finds the differnece in the RA in degrees.
                    d_dec = abs(int(self.jsonDic[target_list[-1]]["Dec"]) - int(
                        self.jsonDic[key]["Dec"]))  # Finds the difference in the Dec in degrees.
                    s = math.sqrt(
                        d_ra ^ 2 + d_dec ^ 2)  # Using the Pythagorean theorem, finding and approximate distance between the targets.
                    if s_o == None or s < s_o:  # If our current distance is less then our current smallest distance.
                        next_target = key  # Setting the next_target which is the key in the dictionary to be our next target value.
                        s_o = s  # Sets our newst smallest distance to be the smallest distance.
            target_list.append(next_target)  # Appends the next target into the sorted target list.
        return target_list

    def running(self):
        """
        This function is called when the self.var variable is set to "Run".
        """
        Button(self.mp, text='Start Over', command=None, bg="gray69").grid(row=4, column=6)  # Disarms restart button.
        self.w.config(bg='green')  # Sets the pause/play button color to green.
        if self.tel.Slewing == False and self.cam.CameraStatus == 2:  # Checking if the Camera and Telescope are Idle.
            self.save_image()
            # Checking if we have already taken images of the target.
            if int(self.jsonDic[self.target_list[self.k]]["ExposureTaken"]) == 0:
                self.JNowConversion()
                # If the current Right Ascension and Declination are within a certain tolerance it will take an image.
                if abs(self.tel.RightAscension - self.TargetRightAscension) < 0.001 and abs(self.tel.Declination - self.TargetDeclination) < 0.002:
                    self.light_image()
                else:  # If not within the RA and Dec tolerances the telescope will slew to the target.
                    self.slew()

    def JNowConversion(self):
        RA  = self.jsonDic[str(self.target_list[self.k])]['RA']
        Dec = self.jsonDic[str(self.target_list[self.k])]['Dec']
        c = SkyCoord(RA, Dec, frame='fk5', unit=(u.hourangle, u.deg), equinox='J2000')
        utcNow = datetime.utcnow()
        t = Time("%s-%s-%s %s:%s:%s" % (
            utcNow.strftime("%Y"), utcNow.strftime("%m"), utcNow.strftime("%d"), utcNow.strftime("%H"),
            utcNow.strftime("%M"), utcNow.strftime("%S")), scale='utc')
        jnow = FK5(equinox=Time(t.jd, format='jd'))
        a = c.transform_to(jnow)
        self.TargetRightAscension = ((a.ra * u.deg).value) / 15.0  # ASCOM needs 0-24 degrees not 0-360 degrees
        self.TargetDeclination = (a.dec * u.deg).value

    def restart(self):
        """
        Let's us go back to the top of our target list
        """
        self.k = 0  # Variable that is iterated across for our targets gets set back to 0.
        self.u = 0  # Variable for exposures gets reset back to 0.
        with open(self.text, 'r') as infile:  # Resets the dictionary which checks for exposures.
            self.jsonDic = json.load(infile)

    def temp_update(self):
        """
        Updates the value for the current temperature of the CCD Camera and will display the temperature in green if
        it is within the temperature tolerance and red if out of the tolerance on the GUI.
        """
        if (self.set_temp - 2) <= self.cam.Temperature <= (self.set_temp + 2):
            self.temp.config(text=round(self.cam.Temperature, 1), font=("Helvetica", 12), bg='green2')
        else:
            self.temp.config(text=round(self.cam.Temperature, 1), font=("Helvetica", 12), bg='red3')

    def value_update(self):
        """
        Updates the Current time (UTC), Right Ascension, and Declination to have the GUI display these values.
        """
        dt = datetime.utcnow().strftime("%H:%M:%S")
        self.phase.config(text=dt, font=("Helvetica", 12), bg='slategrey')
        self.RA = self.tel.RightAscension
        self.RA = round(self.RA, 3)
        self.right.config(text=self.RA, font=("Helvetica", 12), bg='slategrey')
        self.Dec = self.tel.Declination
        self.Dec = round(self.Dec, 3)
        self.Decline.config(text=self.Dec, font=("Helvetica", 12), bg='slategrey')

    def status_update(self):
        """
        Updates the GUI to let us know the status of the CCD camera and the slewing status of the telescope.
        If the telescope is moving the GUI will change it's status to red, and green when not slewing.
        The camera status will display green if there are not errors or red when there is an error or the camera is
        disconnected.
        """
        self.color = self.cam.Filter
        self.fil.config(text=filters_c[self.color], font=("Helvetica", 12), bg='slategrey')
        self.camstat = camstatus[self.cam.CameraStatus]
        if self.cam.CameraStatus == 0 or self.cam.CameraStatus == 1:
            self.camcom.config(text=self.camstat, font=("Helvetica", 12), bg='red3')
        else:
            self.camcom.config(text=self.camstat, font=("Helvetica", 12), bg='Green2')
        if self.tel.Slewing:
            self.scopestat.config(text="Slewing", font=('Helvetica', 12), bg='red3')
        else:
            self.scopestat.config(text="Not Slewing", font=('Helvetica', 12), bg='green2')

    def save_image(self):
        if self.u != 0:  # If we have taken a image ready of the current target we will be able to save it.
            save_name = (self.save + "\\" + self.target_list[self.k] + '_' + str(self.u) + '.fit')
            if os.path.exists(save_name) == False:  # Checking to see if an image with the name exists already.
                self.cam.SaveImage(save_name)  # It did not so we save it as the initial file name.
            else:  # The original file name is already saved in the folder so we save it as a new name
                self.cam.SaveImage(self.save + "\\" + self.target_list[self.k] + '_' + str(self.u) + '(1)' + '.fit')

            # Resetting the number of exposures taken back to 0 and moving on to the next target once the desired number
            # of exposures of a target.
            if self.u == self.jsonDic[self.target_list[self.k]]["NUM"]:
                self.u = 0
                # Appends a value in the target dictionary, so python knows light images
                # were already taken for the curent target.
                self.jsonDic[str(self.target_list[self.k])]["ExposureTaken"] = 1
                if self.k < (len(self.jsonDic.keys()) - 1):
                    self.k = self.k + 1

    def light_image(self):
        """
        Takes a light Image and updates the GUI for how many light images have been taken of a target.
        """
        self.cam.Expose(self.jsonDic[self.target_list[self.k]]["EXP"],
                        1, 11)
        time.sleep(
            0.25)  # This sleep function is here to give the camera enough time to send the signal that it is taking an image.
        # Updating the GUI for the number of light images taken for a target.
        self.u = self.u + 1
        expose = str(self.u) + "/" + str(self.jsonDic[str(self.target_list[self.k])]["NUM"])
        self.light.config(text=expose, font=("Helvetica", 12), bg='slategrey')

    def slew(self):
        # Telescope is going to slew to the target location in the sky.
        self.tel.SlewToCoordinatesAsync(self.TargetRightAscension, self.TargetDeclination)
        # Updating the GUI for the progress through the target list.
        self.target.config(text=self.target_list[self.k], font=("Helvetica", 12), bg='slategrey')
        number_targets = str(self.k + 1) + '/' + str(len(self.target_list))
        self.target_number.config(text=number_targets, font=("Helvetica", 12), bg='slategrey')
        if self.target_list[self.k] != self.target_list[-1]:  # Updating the next target box in our GUI.
            self.next.config(text=self.target_list[(self.k + 1)], font=("Helvetica", 12), bg='slategrey')
        else:  # If we are at the last target in our target list it updates our next target to be None.
            self.next.config(text="None", font=("Helvetica", 12), bg='slategrey')
        # Resetting the number of exposures box in the GUI.
        expose_reset = '0/' + str(self.jsonDic[self.target_list[self.k]]['NUM'])
        self.light.config(text=expose_reset, font=("Helvetica", 12), bg='slategrey')
        return

    def ConnectToTelescope(self, TelescopeName):
        """
        Establishes the link between the telescope and computer, and turns tracking on.
        """
        self.tel = Dispatch(TelescopeName)  # Calls the Telescope and the variable self.tel becomes the telescope name.
        self.tel.Connected = True  # Establishes connection to the Telescope.
        self.tel.Tracking = True  # Turns tracking on the telescope on.
        return self.tel

    def ConnectToCamera(self, CameraName, temp):
        """
        Establishes the link between the CCD camera and computer, and sets the operating temperature.
        """
        self.cam = Dispatch(CameraName)  # Calls the camera and the variable self.cam becomes the camera name.
        self.cam.LinkEnabled = True  # Establishes connection to the camera.
        if self.cam.CanSetTemperature:
            self.cam.TemperatureSetpoint = temp  # Setting the camera operating temperature.
            self.cam.CoolerOn = True  # Turns the Cooler on.
        self.cam.DisableAutoShutdown = True  # Prevents the CCD fom turing off.
        return self.cam


tam = "ASCOM.SiTechDll.Telescope"
cam = 'MaxIm.CCDCamera'
text = "galaxy_targets.txt"
CCDCam = -30
savepath = 'C:\\TelescopeAutomationResearch_DATA'
Automation(cam, tam, text, CCDCam, savepath)
