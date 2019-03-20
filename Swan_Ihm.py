from module_perso.module_graphiques import progress_bar
from module_perso.module_Ihm import *
import os, glob
import json
import pyperclip
from module_perso.module_office import Outlook
import arrow
from time import sleep
import threading
from tkinter import *

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support import ui

import logging
from logging.handlers import RotatingFileHandler

# création de l'objet logger qui va nous servir à écrire dans les logs
log = logging.getLogger()
# on met le niveau du logger à DEBUG, comme ça il écrit tout
log.setLevel(logging.DEBUG)
# création d'un formateur qui va ajouter le temps, le niveau
# de chaque message quand on écrira un message dans le log
formatter = logging.Formatter('%(asctime)s :: %(levelname)s :: %(message)s')
# création d'un handler qui va rediriger une écriture du log vers
# un fichier en mode 'append', avec 1 backup et une taille max de 1Mo
file_handler_debug = RotatingFileHandler('debug.log', 'a', 1000000, 1, encoding="utf-8")
file_handler_info = RotatingFileHandler('info.log', 'a', 1000000, 1, encoding="utf-8")
# on lui met le niveau sur DEBUG, on lui dit qu'il doit utiliser le formateur
# créé précédement et on ajoute ce handler au logger
file_handler_debug.setLevel(logging.WARNING)
file_handler_debug.setFormatter(formatter)
log.addHandler(file_handler_debug)
file_handler_info.setLevel(logging.INFO)
file_handler_info.setFormatter(formatter)
log.addHandler(file_handler_info)
# création d'un second handler qui va rediriger chaque écriture de log
# sur la console
stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
log.addHandler(stream_handler)


def logger(texte="renseigner message", niveau="info"):
    if niveau == "info":
        log.info(texte)
    elif niveau == "debug":
        log.debug(texte)


class Thread(threading.Thread):
    def __init__(self, win, choix_thread, numswan, mailtech):
        threading.Thread.__init__(self)
        self.ns = numswan
        self.choix = choix_thread
        self.win = win
        self.mailtech = mailtech
        self.log = log

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        self.choix(self.ns, self.mailtech)
        self.win.event_generate("<<thread_fini>>")


class Thread_lot(threading.Thread):
    def __init__(self, win, fonction):
        threading.Thread.__init__(self)
        self.win = win
        self.fonction = fonction

    def run(self):
        self.win.event_generate("<<thread_lot_encours>>")
        self.fonction()
        self.win.event_generate("<<thread_lot_fini>>")


class swan_FF(webdriver.Firefox):
    def __init__(self, profil=False):
        # defini l'objet courant comme une instance de firefox utilisant le profil
        profil = self.get_firefox_profile()
        # webdriver.Firefox.__init__(self, 'C:\\Applications\\CAT_VM_Installer\\misc\\FFProfileWeb')
        options = webdriver.FirefoxOptions()
        # options.add_argument('-headless')
        webdriver.Firefox.__init__(self, profil, firefox_options=options)
        # init variable
        self.mail_superviseur_Rennes = data_dict["mail_superviseur_Rennes"]
        self.mail_superviseur_Lannion = data_dict["mail_superviseur_Lannion"]
        self.mail_KLIF = data_dict["mail_KLIF"]
        self.mail_EUROINFO = data_dict["mail_EUROINFO"]
        self.etatvalide = "Title_HMenu_repeater_childRepeater_5_hmenu_linkButtonLevel2_3"
        self.etatdemandeinfo = "Title_HMenu_repeater_childRepeater_5_hmenu_linkButtonLevel2_0"
        self.date = arrow.now().format("DD/MM")
        self.date2 = arrow.now().format("DD/MM/YYYY")
        self.heure = arrow.now().format("HH")
        self.min = arrow.now().format("mm")
        self.ODAGONNEAU = data_dict["acteur_ODAGONNEAU"]
        self.PBRECHET = data_dict["acteur_PBRECHET"]
        self.euroinfo_acteur = data_dict["acteur_euroinfo"]
        # self.rie_acteur = data_dict["acteur_RIE"]
        # self.implicitly_wait(10)
        self.set_page_load_timeout(10)
        self.ouvrir("http://swan.sso.infra.ftgroup/binswan/StartPage.aspx")

    def log(self, log1, log2):
        self.log = log1
        self.log2 = log2

    def get_firefox_profile(self):
        """
        Methode interne qui recupere le profil firefox - ne pas utiliser
        """
        APPDATA = os.getenv('APPDATA')
        FF_PRF_DIR = "%s\\Mozilla\\Firefox\\Profiles\\" % APPDATA
        PATTERN = FF_PRF_DIR + "*default*"
        FF_PRF_DIR_DEFAULT = glob.glob(PATTERN)[0]
        return webdriver.FirefoxProfile(FF_PRF_DIR_DEFAULT)

    def ouvrir(self, url):
        logger("ouverture swan")
        self.get(url)
        WebDriverWait(self, 10).until(expected_conditions.new_window_is_opened)
        sleep(3)

        # recupere la derniere fenetre ouverte (app que l'on vient de lancer)
        app_window = self.window_handles[-1]

        sleep(3)
        # change de fenetre
        self.switch_to.window(app_window)
        WebDriverWait(self, 10).until(expected_conditions.title_is("SWAN - Accueil"))
        # WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.XPATH, "//*[@id='operationSearchButton']")))
        print(self.title)

        ActionChains(self).send_keys(Keys.F6).perform()
        ActionChains(self).send_keys(Keys.F5).perform()
        sleep(1)
        self.maximize_window()
        WebDriverWait(self, 10).until(
            expected_conditions.presence_of_element_located((By.XPATH, "//*[@id='operationSearchButton']")))

    def wait_id(self, locator):
        try:
            # WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.ID, locator)))
            # WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.ID, locator)))
            # WebDriverWait(self, 2).until(expected_conditions.element_located_to_be_selected((By.ID, locator)))
            # WebDriverWait(self, 5).until(expected_conditions.element_located_selection_state_to_be((By.ID, locator)))
            WebDriverWait(self, 5).until(expected_conditions.element_to_be_clickable((By.ID, locator)))
            # sleep(1)
        except WebDriverException:
            logger(str(WebDriverException), "debug")

    def click_by_XPATH(self, locator):
        try:
            # WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.XPATH, locator)))
            # WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.XPATH, locator)))
            WebDriverWait(self, 10).until(expected_conditions.element_to_be_clickable((By.XPATH, locator)))
            # sleep(1)
            self.find_element_by_xpath(locator).click()
        except WebDriverException as e:
            logger(str(e))
            print("erreur selenium : ", e)

    def rech(self, num_swan):
        logger("recherche swan")
        self.find_element_by_id("operationSearchTextBox_Textbox_operationSearchTextBox").clear()
        self.find_element_by_id("operationSearchTextBox_Textbox_operationSearchTextBox").send_keys(num_swan)
        self.find_element_by_id("operationSearchButton").click()
        if self.find_element_by_id("lblErrorMessage").text == "Opération verrouillée":
            self.recherche_swan = False
            self.error = "opération verouillée !!! relancer le traitement aprés suppréssion du vérouillage"
            logger("opération verouillée !!! relancer le traitement aprés suppréssion du vérouillage")

        elif self.find_element_by_id("lblErrorMessage").text == "Numéro de l'opération invalide":
            self.recherche_swan = False
            self.error = "traitement impossible Numéro de l'opération invalide"
            logger("traitement impossible Numéro de l'opération invalide")
        else:
            try:
                WebDriverWait(self, 5).until(
                    expected_conditions.presence_of_element_located(
                        (By.XPATH, "//*[@id='Title_MyBanner_lblTitrepart2']")))
                self.recherche_swan = True
            except WebDriverException as e:
                logger(str(e))
                print("erreur selenium : ", e)

        # assert num_swan in element.text

    def btnmodif(self):
        logger("Modifier")
        self.find_element_by_id("MainContent_btnModify").click()
        WebDriverWait(self, 5).until(
            expected_conditions.text_to_be_present_in_element((By.ID, 'Title_MyBanner_lblTitre'), "Modification"))

    def btnenrquit(self):
        """Enregistrer quitter """
        logger("Enregister quitter")
        try:
            self.wait_id("MainContent_btnSaveAndExit")
            WebDriverWait(self, 10).until(
                expected_conditions.visibility_of_element_located((By.ID, "MainContent_btnSaveAndExit")))
            self.find_element_by_id("MainContent_btnSaveAndExit").click()
            self.wait_id("MainContent_lblFavTitle")
        except WebDriverException as e:
            logger(str(e))
            print("erreur selenium : ", e)

        # self.wait_id("MainContent_lblFavTitle")

    def recup_info(self):
        logger("Recupération info swan")
        self.find_element_by_xpath("//*[@title='commentaires']").click()
        raison_social = self.find_element_by_id("Title_MyBanner_lblComapnyValue").text
        date_deb = self.find_element_by_id("Title_MyBanner_lblDateDebutValue").text
        date_fin = self.find_element_by_id("Title_MyBanner_lblDateFinValue").text
        commentaires = self.find_element_by_id("ul_rptComments").text
        return raison_social, date_deb, date_fin, commentaires

    def changer_etat(self, etat):
        logger("passage état " + etat)
        # clic onglet description
        self.find_element_by_xpath("//*[@title='description']").click()
        self.wait_id("Title_HMenu_repeater_hmenu_hyperLinkLevel1_5")
        self.find_element_by_id("Title_HMenu_repeater_hmenu_hyperLinkLevel1_5").click()
        self.wait_id("Title_HMenu_repeater_childRepeater_5_hmenu_linkButtonLevel2_3")
        self.find_element_by_id(etat).click()

    def activation(self):
        logger("activation")
        # click menu description
        self.find_element_by_xpath("//*[@title='description']").click()
        self.wait_id("Title_HMenu_repeater_childRepeater_0_hmenu_linkButtonLevel2_2")
        # click objet onglet
        self.find_element_by_id("Title_HMenu_repeater_childRepeater_0_hmenu_linkButtonLevel2_2").click()
        self.wait_id("MainContent_btnActivation")
        # click activation
        self.find_element_by_id("MainContent_btnActivation").click()
        sleep(2)

    def commentaires(self, texte):
        logger("modif commentaires")
        # clic onglet description
        self.find_element_by_xpath("//*[@title='description']").click()
        # attente onglet compléments
        self.click_by_XPATH("//*[@title='compléments']")
        WebDriverWait(self, 10).until(
            expected_conditions.visibility_of_element_located((By.XPATH, "//input[contains(@id,'Textbox__2')]")))
        self.find_element_by_xpath("//input[contains(@id,'Textbox__2')]").clear()
        self.find_element_by_xpath("//input[contains(@id,'Textbox__2')]").send_keys(texte)

    def commentaires_OTbacara(self, texte):
        logger("modif commentaires OT bacara")
        self.find_element_by_xpath("//*[@title='description']").click()
        # attente onglet compléments
        self.click_by_XPATH("//*[@title='compléments']")
        # WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.XPATH, "//input[contains(@id,'Textbox__0')]")))
        WebDriverWait(self, 10).until(
            expected_conditions.visibility_of_element_located((By.XPATH, "//input[contains(@id,'Textbox__0')]")))
        self.find_element_by_xpath("//input[contains(@id,'Textbox__0')]").clear()
        self.find_element_by_xpath("//input[contains(@id,'Textbox__0')]").send_keys(texte)

    def modif_acteur(self, acteur):
        logger("modification acteur")
        # click menu description
        self.find_element_by_xpath("//*[@title='description']").click()

        self.click_by_XPATH("//*[@title='acteurs']")
        # sleep(1)
        WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))
        self.wait_id("popUpImageButton2_20")
        self.find_element_by_id("popUpImageButton2_20").click()
        WebDriverWait(self, 10).until(expected_conditions.frame_to_be_available_and_switch_to_it("iframeCommonDialog"))
        # self.switch_to.frame("iframeCommonDialog")
        # sleep(1)
        # WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.ID, "Title_TechnicianListPopupTitle")))
        WebDriverWait(self, 10).until(
            expected_conditions.visibility_of_element_located((By.XPATH, "//a[contains(@href," + acteur + ")]")))

        element = self.find_element_by_xpath("//a[contains(@href," + acteur + ")]")
        position = element.location["y"]
        # self.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        self.execute_script(
            f"window.scrollTo(0, {position});")  # fait defiler la fenetre jusqua la position de l'élément
        # sleep(1)
        self.click_by_XPATH("//a[contains(@href," + acteur + ")]")
        self.switch_to.default_content()
        # sleep(1)
        WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))

    def ot_bacara(self):
        logger("recuperation OT bacara")
        # click menu description
        self.find_element_by_xpath("//*[@title='description']").click()
        self.click_by_XPATH("//*[@title='détails']")
        WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))
        self.wait_id("MainContent_OperationContentPage_hlnkActivationId")
        # WebDriverWait(self, 10).until(expected_conditions.element_to_be_clickable((By.ID, "MainContent_OperationContentPage_hlnkActivationId")))
        # sleep(2)
        self.find_element_by_id("MainContent_OperationContentPage_hlnkActivationId").click()
        self.switch_to.frame("iframeCommonDialog")
        self.wait_id("MainContent_tvlBacara")
        WebDriverWait(self, 5).until(expected_conditions.presence_of_element_located(
            (By.XPATH, "//*[@id='MainContent_tvlBacara']/tbody/tr/td[1]/span")))
        # sleep(2)
        self.bacara = self.find_element_by_xpath("//*[@id='MainContent_tvlBacara']/tbody/tr/td[1]/span").text
        print(self.bacara)
        pyperclip.copy(self.bacara)
        # sleep(2)
        self.find_element_by_xpath("//*[@id='MainContent_btnClose']/span/span/span").click()
        self.switch_to.default_content()
        self.commentaires_OTbacara(self.bacara)
        return self.bacara

    def activation_NFIT(self):
        logger("Activation NFIT")
        # clic onglet Actions
        self.find_element_by_xpath("//*[@title='actions']").click()
        self.click_by_XPATH("//*[@title='créer ticket Océane']")
        self.wait_id("MainContent_OperationContentPage_hlkInputEdsPilot")
        self.find_element_by_id("MainContent_OperationContentPage_hlkInputEdsPilot").click()
        self.switch_to.frame("iframeCommonDialog")
        self.wait_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier")
        self.find_element_by_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier").clear()
        self.find_element_by_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier").send_keys(
            "ATQGAM")
        self.find_element_by_id("MainContent_UcGroupSearch_btnSearch").click()
        self.click_by_XPATH("//*[@title='Choisir']")
        self.switch_to.default_content()
        self.wait_id("MainContent_OperationContentPage_dtStartDate_txtDate_")
        self.find_element_by_class_name("ui-datepicker-trigger").click()
        WebDriverWait(self, 5).until(expected_conditions.element_to_be_clickable(
            (By.CLASS_NAME, "ui-datepicker-days-cell-over.ui-datepicker-today")))
        # sleep(1)
        self.find_element_by_class_name("ui-datepicker-days-cell-over.ui-datepicker-today").click()
        WebDriverWait(self, 5).until(expected_conditions.element_to_be_clickable(
            (By.CLASS_NAME, "ui-datepicker-days-cell-over.ui-datepicker-today")))
        # sleep(1)
        self.wait_id("MainContent_OperationContentPage_ddlTicketOrigin_DrDwList_ddlTicketOrigin")
        ui.Select(self.find_element_by_id(
            "MainContent_OperationContentPage_ddlTicketOrigin_DrDwList_ddlTicketOrigin")).select_by_visible_text(
            "Client")
        ui.Select(self.find_element_by_id(
            "MainContent_OperationContentPage_ddlTicketType_DrDwList_ddlTicketType")).select_by_visible_text(
            "Paramètrage")
        ui.Select(self.find_element_by_id(
            "MainContent_OperationContentPage_ddlTechnicalImpact_DrDwList_ddlTechnicalImpact")).select_by_visible_text(
            "Sans perturbation du service")
        self.click_by_XPATH("//*[@title='Créer ticket']")
        sleep(2)

    def reaffecter_EDS(self, EDS):
        logger("Réaffecter EDS")
        # clic onglet Actions
        self.find_element_by_xpath("//*[@title='actions']").click()
        self.click_by_XPATH("//*[@title='Réaffecter EDS']")
        WebDriverWait(self, 10).until(expected_conditions.frame_to_be_available_and_switch_to_it("iframeCommonDialog"))
        # self.switch_to.frame("iframeCommonDialog")
        self.wait_id("MainContent_lblMainText")
        self.click_by_XPATH("//*[@title='Rechercher un EDS']")
        # sleep(2)
        # self.find_element_by_xpath("//*[@title='Rechercher un EDS']").click()
        self.switch_to.frame("iframeCommonDialog")
        self.wait_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier")
        self.find_element_by_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier").clear()
        self.find_element_by_id("MainContent_UcGroupSearch_txtGroupIdentifier_Textbox_txtGroupIdentifier").send_keys(
            EDS)
        self.find_element_by_id("MainContent_UcGroupSearch_btnSearch").click()
        self.click_by_XPATH("//*[@title='Choisir']")
        self.switch_to.parent_frame()
        self.wait_id("ui-id-1")
        self.find_element_by_id("MainContent_btnValidateButton").click()
        self.switch_to.default_content()

    def objet_mail(self, numswan, info):
        """ info()[0] : raison social
         info[1] : date debut
         info()[2] : date fin """
        objet = "swan " + numswan + " " + info[0] + " pour le " + info[1] + " -> " + info[2] + " : "
        return objet

    def choix_1(self, numswan, tech):
        # qualification -> commentaires 'OD'
        self.btnmodif()
        # self.log2.set("modif commentaire")
        self.commentaires(data_dict["choix_1"]["commentaires"])
        # self.log2.set("modif acteur")
        self.modif_acteur(data_dict["choix_1"]["acteur_choix_1"])
        self.btnenrquit()

    def choix_2(self, numswan, tech):
        # mail demande ressources CP Rennes
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " dem Tech")
        info_swan = self.recup_info()
        self.btnenrquit()
        Outlook(self.mail_superviseur_Rennes, su=self.objet_mail(numswan, info_swan), bd=info_swan[3]).mail()

    def choix_3(self, numswan, tech):
        # mail demande ressources CP Lannion
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " dem Tech")
        info_swan = self.recup_info()
        self.btnenrquit()
        Outlook(self.mail_superviseur_Lannion, su=self.objet_mail(numswan, info_swan), bd=info_swan[3]).mail()

    def choix_4(self, numswan, tech):
        # activation + validation + RDV LA POSTE
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " ressource CP Rennes OK")
        self.activation()
        OT_bacara = self.ot_bacara()
        self.changer_etat(self.etatvalide)
        info_swan = self.recup_info()
        self.btnenrquit()
        Outlook(tech, self.mail_KLIF, self.objet_mail(numswan, info_swan) + OT_bacara,
                info_swan[3],
                info_swan[1], info_swan[2]).rdv()

    def choix_5(self, numswan, tech):
        # activation + validation
        self.btnmodif()
        # activation
        self.activation()
        self.ot_bacara()
        # validation
        self.changer_etat(self.etatvalide)
        self.btnenrquit()

    def choix_6(self, numswan, tech):
        # Validation + RDV EURO INFO BTIP
        self.btnmodif()
        self.modif_acteur(self.euroinfo_acteur)
        info_swan = self.recup_info()
        self.changer_etat(self.etatvalide)
        self.btnenrquit()
        Outlook(self.mail_EUROINFO, self.mail_KLIF, self.objet_mail(numswan, info_swan),
                info_swan[3], info_swan[1], info_swan[2]).rdv()

    def choix_7(self, numswan, tech):
        # Creation  RDV
        info_swan = self.recup_info()
        Outlook("", self.mail_KLIF, self.objet_mail(numswan, info_swan), info_swan[3],
                info_swan[1], info_swan[2]).rdv()

    def choix_8(self, numswan, tech):
        # Cretion Email
        info_swan = self.recup_info()
        Outlook("", su=self.objet_mail(numswan, info_swan), bd=info_swan[3]).mail()

    def choix_9(self, numswan, tech):
        # Activation NFIT
        self.btnmodif()
        self.activation_NFIT()
        self.changer_etat(self.etatvalide)
        self.btnenrquit()

    def choix_10(self, numswan, tech):
        # test MCS SITA
        self.btnmodif()
        self.commentaires("test mcs SITA")
        # self.modif_acteur(self.ODAGONNEAU)
        self.activation()
        self.ot_bacara()
        info_swan = self.recup_info()
        self.changer_etat(self.etatvalide)
        self.btnenrquit()
        Outlook(data_dict["SITA"]["mail"], self.mail_KLIF,
                "test MCS site WS-ARUN-TOWN1 / RUNDD / CD043B9FF4", data_dict["SITA"]["texte"],
                info_swan[1], info_swan[2]).rdv()

    def debug(self, numswan, tech):
        # self.btnmodif()
        info_swan = self.recup_info()
        OT_bacara = self.ot_bacara()
        Outlook(tech, self.mail_KLIF, self.objet_mail(numswan, info_swan) + OT_bacara,
                info_swan[3], info_swan[1], info_swan[2]).rdv()
        # rdv(dest=data_dict["SITA"]["mail"], dest_op=self.mail_KLIF, sujet="test MCS site WS-ARUN-TOWN1 / RUNDD / CD043B9FF4", corps=data_dict["SITA"]["texte"], date_deb=info_swan[1], date_fin=info_swan[2])
        # self.reaffecter_EDS("ATQCHR")
        # self.commentaires("OD")
        # self.changer_etat(self.etatvalide)
        # self.modif_acteur(data_dict["choix_1"]["acteur_choix_1"])
        # self.btnenrquit()
        # info_swan = self.recup_info()
        # ot_bacara = self.ot_bacara()
        # Outlook(tech, self.mail_KLIF, self.objet_mail(numswan, info_swan) + "ot_bacara", info_swan[3], info_swan[1],info_swan[2]).rdv()
        # Outlook("", su=self.objet_mail(numswan, info_swan), bd=info_swan[3]).mail()
        # self.changer_etat(self.etatvalide)
        # self.btnenrquit()
        # objet = "swan " + numswan + " " + info_swan[0] + " pour le " + info_swan[1] + "->" + info_swan[2]
        # rdv("", self.mail_KLIF, objet, info_swan[3], info_swan[1], info_swan[2])
        # sleep(2)


class Swan_IHM(interface_IHM):
    """  robot Swan"""

    def __init__(self, action="json\\swan_ihm.json", repjson="json", titre=50 * " " + "ROBOT SWAN"):
        self.liste_action_swan = [swan.choix_1, swan.choix_2, swan.choix_3, swan.choix_4, swan.choix_5, swan.choix_6,
                                  swan.choix_7, swan.choix_8, swan.choix_9, swan.choix_10, swan.debug]
        self.liste_tech = [tech for tech in data_dict["techs_la_poste"].keys()]
        # print(self.liste_tech)
        self.mailtech = ""
        """ constructeur de l'IHM """
        super().__init__(action, repjson, titre)
        self.root.bind('<<fincombox1>>', self.action_widget_combolist1)
        self.root.bind('<<thread_fini>>', self.nettoyagelog)

    def nettoyagelog(self, e):
        print("nettoyage log fin de thread")
        self.sv.set("")
        self.barre1.barre.destroy()
        self.mailtech = ""
        pass

    def trait_fichier_lot(self):
        import xlrd
        self.fich_traitlot = filedialog.askopenfilename(filetypes=[("Excel", "*.xls")],
                                                        initialdir="C:\Applications\Robotswan\\trait_lot",
                                                        title="fichier traitement par lot")
        classeur = xlrd.open_workbook(self.fich_traitlot)
        feuill1 = classeur.sheet_by_name(classeur.sheet_names()[0])  # nom feuill1
        self.traitement_lot = TRUE
        self.fic_trait.set(self.fich_traitlot)
        self.entree2.update()
        self.sv2.set("Traitement par lot en cours .....")  # affichage du traitement en cours
        self.lab2.update()
        barre2 = progress_bar(self.fr_log, feuill1.nrows - 1, 300)
        barre2.barre.grid(row=3, column=0)
        barre2.lab_pourcent.configure(bg="ORANGE", fg="blue")
        barre2.lab_pourcent.grid(row=3, column=1)
        for ligne_index in feuill1.col(0)[1:]:
            self.choix = str(self.liste.get(self.liste.curselection()))
            self.ref_saisie = ligne_index.value
            self.ind_action = self.liste.curselection()[0]
            self.trait_unitaire()
            print(ligne_index.value)
            self.t.join()
            barre2.progression()
        self.traitement_lot = FALSE
        self.fic_trait.set("")
        self.entree2.update()
        self.sv2.set("")
        self.lab2.update()
        barre2.barre.destroy()
        barre2.lab_pourcent.destroy()
        self.liste.select_clear(0, 'end')

    def thread_traitement_lot(self):
        t2 = Thread_lot(self.root, self.trait_fichier_lot)
        t2.start()

    def ouvrir_modele_traitlot(self):
        subprocess.call(["C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE",
                         "C:\Applications\Robotswan\\trait_lot\modele_trait_lot_swan_1.xls"])

    def action(self):
        self.ref_saisie = self.entree.get()
        self.choix = str(self.liste.get(self.liste.curselection()))
        print(self.choix)
        logger(f"{self.ref_saisie}  --  {self.choix}")
        self.ind_action = self.liste.curselection()[0]
        if self.liste.curselection() == ():  # verification action sélectionnée
            self.sv.set("veuillez choisir une action !!! ")
        elif self.entree.get() == "":
            self.sv.set("Veuiller saisir une réference !!!!")  # affichage msg erreur dans bandeau
        elif self.liste.curselection()[0] == 3:
            self.sv.set(f"traitement en cours : {self.ref_saisie} -> {self.choix}")
            self.combo1 = widget_ComboBox(self.root, self.liste_tech, 1, "choix tech", "selectionner un technicien")
        else:
            # realiser les actions en fonctions de l'action selectionnée
            self.trait_unitaire()

    def trait_unitaire(self):
        swan.rech(self.ref_saisie)
        if swan.recherche_swan:
            self.sv.set(f"traitement en cours : {self.ref_saisie} -> {self.choix}")
            self.barre1 = progress_bar(self.fr_log, 10, 200, "yellow", "indeterminate")
            self.barre1.barre.grid(row=1, column=0)
            self.barre1.progression()
            self.fr_log.update()
            self.t = Thread(self.root, self.liste_action_swan[self.ind_action], self.ref_saisie, self.mailtech)
            self.t.start()
        else:
            self.sv.set(swan.error)
            self.lab1.update()
            self.liste.select_clear(0, 'end')
        if not self.traitement_lot:
            self.liste.select_clear(0, 'end')

    def action_widget_combolist1(self, e):
        print(data_dict["techs_la_poste"][self.combo1.resultats])
        self.mailtech = data_dict["techs_la_poste"][self.combo1.resultats]
        self.combo1.fin()
        swan.rech(self.ref_saisie)
        self.trait_unitaire()

    def quitter(self):
        swan.quit()
        self.root.quit()


if __name__ == "__main__":
    # recuperation des données du fichier swan.json dans un dictionnaire data_dict
    with open("json\swan.json") as swan_data:
        data_dict = json.load(swan_data)

    swan = swan_FF()
    app = Swan_IHM()
