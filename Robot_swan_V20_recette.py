# coding: utf-8
import logging
import os, glob
from selenium.common.exceptions import TimeoutException
from logging.handlers import RotatingFileHandler
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from module_perso.module_graphiques import progress_bar
from selenium.common.exceptions import WebDriverException

from selenium.webdriver.support.ui import Select
# import selenium.common.exceptions as se
import threading

# from selenium.webdriver.support.expected_conditions import presence_of_element_located
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


def rdv(dest, dest_op, sujet, corps, date_deb, date_fin):
    logger("Invit outlook")
    rdv_init = Outlook(dest, dest_op, sujet, corps, date_deb, date_fin)
    rdv_init.rdv()


def email(dest, sujet, corps):
    logger("mail outlook")
    mail_init = Outlook(dest, su=sujet, bd=corps)
    mail_init.mail()
    
    


def logger(texte="renseigner message", niveau="info"):
    if niveau == "info":
        log.info(texte)
    elif niveau == "debug":
        log.debug(texte)


class swan(webdriver.Firefox):
    def __init__(self, profil=False):
        # defini l'objet courant comme une instance de firefox utilisant le profil
        profil = self.get_firefox_profile()
        # webdriver.Firefox.__init__(self, 'C:\\Applications\\CAT_VM_Installer\\misc\\FFProfileWeb')
        options = webdriver.FirefoxOptions()
        #options.add_argument('-headless')
        webdriver.Firefox.__init__(self, profil,firefox_options=options)
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
        #self.rie_acteur = data_dict["acteur_RIE"]
        #self.implicitly_wait(10)
        self.set_page_load_timeout(10)
        self.ouvrir("http://swan.sso.infra.ftgroup/binswan/StartPage.aspx")


    #
    def get_firefox_profile(self):
        """
        Methode interne qui recupere le profil firefox - ne pas utiliser
        """
        APPDATA = os.getenv('APPDATA')
        FF_PRF_DIR = "%s\\Mozilla\\Firefox\\Profiles\\" % APPDATA
        PATTERN = FF_PRF_DIR + "*default*"
        FF_PRF_DIR_DEFAULT = glob.glob(PATTERN)[0]
        return webdriver.FirefoxProfile(FF_PRF_DIR_DEFAULT)

    def ouvrir(self,url):
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
        #WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.XPATH, "//*[@id='operationSearchButton']")))
        print(self.title)

        ActionChains(self).send_keys(Keys.F6).perform()
        ActionChains(self).send_keys(Keys.F5).perform()
        sleep(1)
        self.maximize_window()
        WebDriverWait(self, 10).until(
            expected_conditions.presence_of_element_located((By.XPATH, "//*[@id='operationSearchButton']")))

    def wait_id(self, locator):
        try:
            #WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.ID, locator)))
            #WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.ID, locator)))
            #WebDriverWait(self, 2).until(expected_conditions.element_located_to_be_selected((By.ID, locator)))
            #WebDriverWait(self, 5).until(expected_conditions.element_located_selection_state_to_be((By.ID, locator)))
            WebDriverWait(self, 5).until(expected_conditions.element_to_be_clickable((By.ID, locator)))
            #sleep(1)
        except WebDriverException:
            logger(str(WebDriverException), "debug")

    def click_by_XPATH(self, locator):
        try:
            #WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.XPATH, locator)))
            #WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.XPATH, locator)))
            WebDriverWait(self, 10).until(expected_conditions.element_to_be_clickable((By.XPATH, locator)))
            #sleep(1)
            self.find_element_by_xpath(locator).click()
        except WebDriverException:
            logger(str(WebDriverException), "debug")

    def rech(self, num_swan):
        logger("recherche swan")
        self.find_element_by_id("operationSearchTextBox_Textbox_operationSearchTextBox").clear()
        self.find_element_by_id("operationSearchTextBox_Textbox_operationSearchTextBox").send_keys(num_swan)
        self.find_element_by_id("operationSearchButton").click()
        if self.find_element_by_id("lblErrorMessage").text == "Opération verrouillée":
            self.recherche_swan = FALSE
            self.error = "opération verouillée !!! relancer le traitement aprés suppréssion du vérouillage"
            logger("opération verouillée !!! relancer le traitement aprés suppréssion du vérouillage")
        elif self.find_element_by_id("lblErrorMessage").text == "Numéro de l'opération invalide":
            self.recherche_swan = FALSE
            self.error = "traitement impossible Numéro de l'opération invalide"
            logger("traitement impossible Numéro de l'opération invalide")
        else:
            try:
                WebDriverWait(self, 5).until(
                    expected_conditions.presence_of_element_located(
                        (By.XPATH, "//*[@id='Title_MyBanner_lblTitrepart2']")))
                self.recherche_swan = TRUE
            except WebDriverException as e:
                logger(str(e), "debug")
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
            WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))
            self.find_element_by_id("MainContent_btnSaveAndExit").click()
            self.wait_id("MainContent_lblFavTitle")
        except WebDriverException as e:
            logger(str(e), "debug")
            print("erreur selenium : ", e)

        #self.wait_id("MainContent_lblFavTitle")

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
        #WebDriverWait(self, 10).until(expected_conditions.presence_of_element_located((By.XPATH, "//input[contains(@id,'Textbox__0')]")))
        WebDriverWait(self, 10).until(
            expected_conditions.visibility_of_element_located((By.XPATH, "//input[contains(@id,'Textbox__0')]")))
        self.find_element_by_xpath("//input[contains(@id,'Textbox__0')]").clear()
        self.find_element_by_xpath("//input[contains(@id,'Textbox__0')]").send_keys(texte)

    def modif_acteur(self, acteur):
        logger("modification acteur")
        # click menu description
        self.find_element_by_xpath("//*[@title='description']").click()

        self.click_by_XPATH("//*[@title='acteurs']")
        #sleep(1)
        WebDriverWait(self,10).until(expected_conditions.invisibility_of_element_located((By.ID,"disablingDiv")))
        self.wait_id("popUpImageButton2_20")
        self.find_element_by_id("popUpImageButton2_20").click()
        WebDriverWait(self, 10).until(expected_conditions.frame_to_be_available_and_switch_to_it("iframeCommonDialog"))
        #self.switch_to.frame("iframeCommonDialog")
        # sleep(1)
        #WebDriverWait(self, 10).until(expected_conditions.visibility_of_element_located((By.ID, "Title_TechnicianListPopupTitle")))
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
        #sleep(1)
        WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))

    def OT_bacara(self):
        logger("recuperation OT bacara")
        # click menu description
        self.find_element_by_xpath("//*[@title='description']").click()
        self.click_by_XPATH("//*[@title='détails']")
        WebDriverWait(self, 10).until(expected_conditions.invisibility_of_element_located((By.ID, "disablingDiv")))
        self.wait_id("MainContent_OperationContentPage_hlnkActivationId")
        #WebDriverWait(self, 10).until(expected_conditions.element_to_be_clickable((By.ID, "MainContent_OperationContentPage_hlnkActivationId")))
        #sleep(2)
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

    def Activation_NFIT(self):
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
        WebDriverWait(self,5).until(expected_conditions.element_to_be_clickable((By.CLASS_NAME,"ui-datepicker-days-cell-over.ui-datepicker-today")))
        #sleep(1)
        self.find_element_by_class_name("ui-datepicker-days-cell-over.ui-datepicker-today").click()
        WebDriverWait(self, 5).until(expected_conditions.element_to_be_clickable(
            (By.CLASS_NAME, "ui-datepicker-days-cell-over.ui-datepicker-today")))
        #sleep(1)
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

    def Reaffecter_EDS(self, EDS):
        logger("Réaffecter EDS")
        # clic onglet Actions
        self.find_element_by_xpath("//*[@title='actions']").click()
        self.click_by_XPATH("//*[@title='Réaffecter EDS']")
        WebDriverWait(self, 10).until(expected_conditions.frame_to_be_available_and_switch_to_it("iframeCommonDialog"))
        # self.switch_to.frame("iframeCommonDialog")
        self.wait_id("MainContent_lblMainText")
        self.click_by_XPATH("//*[@title='Rechercher un EDS']")
        #sleep(2)
        #self.find_element_by_xpath("//*[@title='Rechercher un EDS']").click()
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
        self.commentaires(data_dict["choix_1"]["commentaires"])
        self.modif_acteur(data_dict["choix_1"]["acteur_choix_1"])
        self.btnenrquit()

    def choix_2(self, numswan, tech):
        # mail demande ressources CP Rennes
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " dem Tech")
        info_swan = self.recup_info()
        self.btnenrquit()
        email(self.mail_superviseur_Rennes, self.objet_mail(numswan, info_swan), info_swan[3])

    def choix_3(self, numswan, tech):
        # mail demande ressources CP Lannion
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " dem Tech")
        info_swan = self.recup_info()
        self.btnenrquit()
        email(self.mail_superviseur_Lannion, self.objet_mail(numswan, info_swan), info_swan[3])

    def choix_4(self, numswan, tech):
        # activation + validation + RDV LA POSTE
        self.btnmodif()
        self.commentaires(data_dict["choix_1"]["commentaires"] + " " + self.date + " ressource CP Rennes OK")
        self.activation()
        OT_bacara = self.OT_bacara()
        self.changer_etat(self.etatvalide)
        info_swan = self.recup_info()
        self.btnenrquit()
        rdv(dest=tech, dest_op=self.mail_KLIF, sujet=self.objet_mail(numswan, info_swan) + OT_bacara,
            corps=info_swan[3],
            date_deb=info_swan[1], date_fin=info_swan[2])

    def choix_5(self, numswan, tech):
        # activation + validation
        self.btnmodif()
        # activation
        self.activation()
        self.OT_bacara()
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
        rdv(dest=self.mail_EUROINFO, dest_op=self.mail_KLIF, sujet=self.objet_mail(numswan, info_swan),
            corps=info_swan[3], date_deb=info_swan[1], date_fin=info_swan[2])

    def choix_7(self, numswan, tech):
        # Creation  RDV
        info_swan = self.recup_info()
        rdv(dest="", dest_op=self.mail_KLIF, sujet=self.objet_mail(numswan, info_swan), corps=info_swan[3],
            date_deb=info_swan[1], date_fin=info_swan[2])

    def choix_8(self, numswan, tech):
        # Cretion Email
        info_swan = self.recup_info()
        email("", self.objet_mail(numswan, info_swan), info_swan[3])
    
    def choix_9(self, numswan, tech):
        # Activation NFIT
        self.btnmodif()
        self.Activation_NFIT()
        self.changer_etat(self.etatvalide)
        self.btnenrquit()

    def choix_10(self, numswan, tech):
         #test MCS SITA
        self.btnmodif()
        self.commentaires("test mcs SITA")
        #self.modif_acteur(self.ODAGONNEAU)
        self.activation()
        self.OT_bacara()
        info_swan = self.recup_info()
        self.changer_etat(self.etatvalide)
        self.btnenrquit()
        rdv(dest=data_dict["SITA"]["mail"], dest_op=self.mail_KLIF, sujet="test MCS site WS-ARUN-TOWN1 / RUNDD / CD043B9FF4", corps=data_dict["SITA"]["texte"], date_deb=info_swan[1], date_fin=info_swan[2])

    def debug(self, numswan, tech):
        #self.btnmodif()
        info_swan = self.recup_info()
        OT_bacara = self.OT_bacara()
        rdv(dest=tech, dest_op=self.mail_KLIF, sujet=self.objet_mail(numswan, info_swan) + OT_bacara,
            corps=info_swan[3],
            date_deb=info_swan[1], date_fin=info_swan[2])
        #rdv(dest=data_dict["SITA"]["mail"], dest_op=self.mail_KLIF, sujet="test MCS site WS-ARUN-TOWN1 / RUNDD / CD043B9FF4", corps=data_dict["SITA"]["texte"], date_deb=info_swan[1], date_fin=info_swan[2])
        # self.Reaffecter_EDS("ATQCHR")
        # self.commentaires("OD")
        #self.changer_etat(self.etatvalide)
        # self.modif_acteur(data_dict["choix_1"]["acteur_choix_1"])
        #self.btnenrquit()
        # info_swan = self.recup_info()
        # OT_bacara = self.OT_bacara()
        # rdv(tech, self.mail_KLIF, self.objet_mail(numswan, info_swan) + "OT_bacara", info_swan[3], info_swan[1],info_swan[2])
        # email("", self.objet_mail(numswan, info_swan), info_swan[3])
        # self.changer_etat(self.etatvalide)
        # self.btnenrquit()
        # objet = "swan " + numswan + " " + info_swan[0] + " pour le " + info_swan[1] + "->" + info_swan[2]
        # rdv("", self.mail_KLIF, objet, info_swan[3], info_swan[1], info_swan[2])
        # sleep(2)


class Thread(threading.Thread):

    def __init__(self, win, choix_thread, numswan, mailtech):
        threading.Thread.__init__(self)
        self.ns = numswan
        self.choix = choix_thread
        self.win = win
        self.mailtech = mailtech

    def run(self):
        # self.threadencours = TRUE
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


##########################################################################################################################


class interface(object):
    def __init__(self):

        """construteur de la fenetre principal"""
        # swan = swan()
        # swan.ouvrir()
        self.root = Tk()
        self.root.title("                            ROBOT SWAN         V2.0")
        # self.root.geometry("675x430")
        self.menu()
        self.traitement_lot = FALSE
        self.mailtech = ""
        self.choix_utilisateur = dict([
            (f' qualification -> commentaires : {data_dict["choix_1"]["commentaires"]}', swan.choix_1),
            (" mail  demanderessources CP Rennes", swan.choix_2),
            (" mail demande ressources CP Lannion", swan.choix_3),
            (" activation + validation + RDV LA POSTE", swan.choix_4),
            (" activation + validation", swan.choix_5),
            (" validation + RDV EUROINFO", swan.choix_6),
            (" Creation RDV", swan.choix_7),
            (" Envoi_mail", swan.choix_8),
            (" Activation NFIT", swan.choix_9),
            (" test MCS SITA", swan.choix_10),
            (" Debug fonction", swan.debug)
        ])

        fr_bas = Frame(self.root)
        fr_bas.pack(fill=BOTH, side=BOTTOM)
        fr_gauche = Frame(self.root)
        fr_gauche.pack(fill=BOTH, side=LEFT)
        fr_droite = Frame(self.root)
        fr_droite.pack(fill=BOTH, side=RIGHT)
        self.fr_log = Frame(fr_bas, relief=SUNKEN, bd=5, bg="ORANGE")
        self.fr_log.pack(side=TOP, expand=1, pady=5)

        # frame numero swan
        fr_gauche_haut = Frame(fr_gauche, bd=8, relief=RAISED)
        fr_gauche_haut.pack(side=TOP, padx=10, pady=10)
        Label(fr_gauche_haut, text="N° de Swan à Traiter", font="arial 10 bold").grid(row=0, column=0)
        self.entree = Entry(fr_gauche_haut, font="arial 10 bold", width=10, bg="ORANGE")
        self.entree.grid(row=1, column=0, pady=5)

        # frame traitement par lot
        fr_gauche_bas = Frame(fr_gauche, bd=8, relief=RAISED)
        fr_gauche_bas.pack(side=BOTTOM, padx=5, pady=5)
        Label(fr_gauche_bas, text="Traitement par lot", font="arial 10 bold").grid(row=0, columnspan=2)
        Label(fr_gauche_bas, text="Selection fichier d'import", ).grid(row=1, column=0)
        self.fic_trait = StringVar()
        self.fic_trait.set("")
        self.entree2 = Entry(fr_gauche_bas, width=35, bg="ORANGE", textvariable=self.fic_trait)
        self.entree2.grid(row=2, columnspan=2, pady=5)
        Button(fr_gauche_bas, text="Parcourir", command=self.thread_traitement_lot, height=1, width=10,
               relief=RAISED,
               bd=4).grid(row=1,
                          column=1, padx=5, pady=5
                          )
        Button(fr_gauche_bas, text="Modele", command=self.ouvrir_modele_traitlot, height=1, width=10,
               relief=RAISED, bd=4).grid(row=3,
                                         columnspan=2,
                                         padx=5,
                                         pady=5)

        # liste choix
        # frame selection action
        fr_d = Frame(fr_droite, bd=4, relief=RAISED)
        fr_d.pack(side=RIGHT, padx=10)
        liste_action = Variable(fr_d, (f' qualification -> commentaires : {data_dict["choix_1"]["commentaires"]}',
                                       " mail  demanderessources CP Rennes",
                                       " mail demande ressources CP Lannion",
                                       " test MCS SITA",   
                                       " activation + validation + RDV LA POSTE", " activation + validation",
                                       " validation + RDV EUROINFO", " Creation RDV", " Envoi_mail",
                                       " Activation NFIT", " Debug fonction"))
        self.liste = Listbox(fr_d, listvariable=liste_action, selectmode="single", selectbackground="ORANGE",
                             height=10, width=38, activestyle="none")
        self.liste.grid(row=0, column=0, padx=5)
        Button(fr_d, text="OK", command=self.action).grid(row=1, column=0, pady=10)

        # bouton quitter
        Button(fr_bas, text="Quitter", command=self.quitter, height=1, width=20).pack(side=BOTTOM, pady=5)

        # bandeaux log
        self.sv = StringVar()
        self.lab1 = Label(self.fr_log, textvariable=self.sv, bg="ORANGE")
        self.lab1.grid(row=0, pady=5)
        self.lab3 = Label(self.fr_log, bg="ORANGE",
                          text="                                                                                                                 ")
        self.lab3.grid(row=1, pady=5)

        self.sv2 = StringVar()
        self.lab2 = Label(self.fr_log, textvariable=self.sv2, fg="blue", bg="ORANGE")
        self.lab2.grid(row=2, pady=5)
        self.lab4 = Label(self.fr_log, text="", bg="ORANGE")
        self.lab4.grid(row=3, pady=5)

        self.root.protocol("WM_DELETE_WINDOW", self.quitter)
        self.root.bind('<<thread_fini>>', self.nettoyagelog1)
        self.root.mainloop()

    def menu(self):
        """"gestion menu fen root"""
        self.men = Menu(self.root)
        self.men_admin = Menu(self.men, tearoff=0)
        self.men_admin.add_command(label="Modif fichier swan.json", command=self.ouvrir_swanjson)
        self.men.add_cascade(label="Administration", menu=self.men_admin)
        self.root.config(menu=self.men)

    def ouvrir_swanjson(self):
        self.sv2.set("modif fichier swan.json en cours .....")  # affichage du traitement en cours
        self.lab2.update()
        subprocess.call(["notepad.exe", "c:\Applications\Robotswan\json\swan.json"])
        self.sv2.set("")
        self.lab2.update()
        showwarning("Alerte     ROBOT SWAN    ",
                    message="Attention pour prendre en compte les modifications veuillez redémarrer l'application")
        self.quitter()

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
            self.CurSel(ligne_index.value)
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
        self.choix = str(self.liste.get(self.liste.curselection()))
        if self.choix == " activation + validation + RDV LA POSTE":
            self.choix_tech()
        #   numeroswan = self.entree.get()
        else:
            numswan = self.entree.get()
            self.CurSel(numswan)

    def CurSel(self, numswan):
        # numswan=self.entree.get()
        print(self.choix)
        if self.choix is NONE:
            self.choix = str(self.liste.get(self.liste.curselection()))

        logger(f"{numswan}  --  {self.choix}")

        swan.rech(numswan)

        if swan.recherche_swan:
            self.sv.set(f"traitement en cours : {numswan} -> {self.choix}")  # affichage du traitement en cours
            self.lab1.update()
            # if not self.traitement_lot:
            self.barre1 = progress_bar(self.fr_log, 10, 200, "yellow", "indeterminate")
            self.barre1.barre.grid(row=1, column=0)
            self.barre1.progression()
            self.fr_log.update()
            self.t = Thread(self.root, self.choix_utilisateur[self.choix], numswan, self.mailtech)
            self.t.start()
            # else:
            # self.alert()
            numswan = ""
            if not self.traitement_lot:
                self.liste.select_clear(0, 'end')

        else:
            self.sv.set(swan.error)
            self.lab1.update()
            self.liste.select_clear(0, 'end')

    def choix_tech(self):
        liste_tech = [tech for tech in data_dict["techs_la_poste"].keys()]
        self.fen_choix_tech = Toplevel(self.root)
        self.fen_choix_tech.title("choix tech")
        self.fen_choix_tech.geometry("+900+250")
        lab = Label(self.fen_choix_tech, text="selectionnez le tech à affecter au swan")
        lab.pack(padx=5, pady=5)
        self.liste_choix_tech = Listbox(self.fen_choix_tech, selectbackground="ORANGE", relief=RAISED,
                                        activestyle="none", bd=7)
        for index in range(len(liste_tech)):
            self.liste_choix_tech.insert(index + 1, liste_tech[index])
        self.liste_choix_tech.pack(fill=X, expand=YES, padx=5, pady=5)
        Button(self.fen_choix_tech, text="OK", command=self.action2).pack()

    def action2(self):
        self.mailtech = data_dict["techs_la_poste"][self.liste_choix_tech.get(self.liste_choix_tech.curselection())]
        numswan = self.entree.get()
        # self.CurSel(numeroswan,self.mailtech)
        self.fen_choix_tech.destroy()
        self.CurSel(numswan)

    def nettoyagelog1(self, e):
        print("nettoyage log fin de thread")
        self.sv.set("")
        self.lab1.update()
        # if not self.traitement_lot:
        self.barre1.barre.destroy()
        self.mailtech = ""

    def quitter(self):
        swan.quit()
        self.root.quit()

    def alert(self):
        showinfo("alerte", "Menu en construction")


if __name__ == "__main__":
    from tkinter import *
    from tkinter.messagebox import *
    import subprocess
    from time import sleep
    from selenium.webdriver.support import ui
    from module_perso.module_office import Outlook
    from datetime import datetime
    import arrow
    import pyperclip
    import json
    from tkinter import filedialog

    date = datetime.now()
    # recuperation des données du fichier swan.json dans un dictionnaire data_dict
    with open("json\swan.json") as swan_data:
        data_dict = json.load(swan_data)
    swan = swan()
    robotswan = interface()
    # swan=swan()
    # swan.ouvrir()
    # swan.rech("T807046482")
    # swan.btnmodif()
    # swan.btnenrquit()
    # print(swan.recup_info())
    # swan.ferme()
    
