"""
HealthyW8 -- 

Gamebus Campaign Checker Tool

author: Serge Autexier

copyright: DFKI GmbH 2026 

License: Apache License 2.0 

"""

import os
import logging
import argparse
import math
import pandas as pd
import re
import language_tool_python
from language_tool_python.utils import classify_matches, TextStatus

LOG_NAME = "Gamebus-Campaign-Checker.log"

# errortypes 

CONSISTENCY = "consistency"
VISUALIZATIONINTERN = "visualizationintern"
REACHABILITY = "reachability"
TARGETPOINTSREACHABLE = "targetpointsreachable"
SECRETS = "secrets"
SPELLCHECKER = "spellchecker"
TTMSTRUCTURE = "ttm"
    
class CampaignChecker:
    gc:dict = {}
    errors:dict = {}
    exportfilename:str = ""
    savefilename:str = ""
    
    def __init__(self, filename):
        self.gc = self.load_campaign(filename)
        self.errors = {CONSISTENCY:[],VISUALIZATIONINTERN:[],
                        REACHABILITY:[],TARGETPOINTSREACHABLE:[],
                        SECRETS:[],SPELLCHECKER:[],
                        TTMSTRUCTURE:[]}
        if (os.path.dirname(filename)==""):
            dirname = "."
        else: 
            dirname = os.path.basename(filename)
        self.exportfilename = f"{dirname}{os.path.sep}Errors-{os.path.basename(filename)}"
        self.savefilename = f"{dirname}{os.path.sep}Save-{os.path.basename(filename)}"
        
    def addError(self,type,errormessage,visualization,challenge):
        self.errors[type].append({'error':errormessage,'visualization':visualization,'challenge':challenge})
        
    def checkResult(self,type):
        if (self.errors[type]==[]): 
            return "Passed"
        else: 
            return "Failed"
    
    def errorsToLog(self):
        for kind in [CONSISTENCY,VISUALIZATIONINTERN,REACHABILITY,TARGETPOINTSREACHABLE,SECRETS,SPELLCHECKER]:
            for localerror in self.errors[kind]:
                logging.warning(f"{kind} error: visualization '{localerror['visualization']['description']}', challenge '{localerror['challenge']['name']}': {localerror['error']} (EDIT HERE: {self.getChallengeEditURL(localerror['visualization'],localerror['challenge'])})")

    def errorsToExcel(self):  
        content = {'Kind' : [], 'Visualization':[],'Challenge':[],'Error':[],'URL':[]}
        entries = 0
        for list in content:
            entries = entries + len(content[list])
        if (entries>0):    
            for kind in [CONSISTENCY,VISUALIZATIONINTERN,REACHABILITY,TARGETPOINTSREACHABLE,SECRETS,SPELLCHECKER,TTMSTRUCTURE]:
                for localerror in self.errors[kind]:
                    content['Kind'].append(kind)
                    content['Visualization'].append(localerror['visualization']['description'])
                    content['Challenge'].append(localerror['challenge']['name'])
                    content['Error'].append(localerror['error'])
                    content['URL'].append(self.getChallengeEditURL(localerror['visualization'],localerror['challenge']))
            df = pd.DataFrame(content)
            if os.path.exists(self.exportfilename):
                os.remove(self.exportfilename)
            writer = pd.ExcelWriter(self.exportfilename,engine='xlsxwriter')
            df.to_excel(writer,sheet_name="Errors",startrow=0,startcol=0,index=False)
            worksheet = writer.sheets['Errors']
            (max_row, max_col) = df.shape
            # Create a list of column headers, to use in add_table().
            column_settings = [{'header': column} for column in df.columns]
            # Add the Excel table structure. Pandas will add the data.
            worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
            # Make the columns wider for clarity.
            worksheet.autofit()
            writer.close()
            logging.info(f"Errors are in {self.exportfilename}")
        else:
            if os.path.exists(self.exportfilename):
                os.remove(self.exportfilename)
            logging.info(f"No Errors to export.")
    
    campaignsheetnames = ["campaigns","categories","labels","groups","waves","visualizations","rewards","challenges","tasks","challengerewards","lotteries","remindertemplates","wearableconfigs"]
    
    def load_campaign(self,filename):
        gc = {}
        for s in self.campaignsheetnames:
            gc[s] = pd.read_excel(filename,sheet_name=s)
        return gc

    def campaignToExcel(self):
        writer = pd.ExcelWriter(self.savefilename, engine='xlsxwriter')
        for sheetname in self.campaignsheetnames:
            self.gc[sheetname].to_excel(writer,sheet_name=sheetname,engine='xlsxwriter')
            worksheet = writer.sheets[sheetname]
            worksheet.autofit()
        writer.close()
        logging.info(f"Saved campaign in {self.savefilename}")
    
            
    """"
    ====================================================

    Visualization

    ====================================================
    """

    def getVisualizations(self):
        return self.gc["visualizations"]

    def getVisualizationChallengesbyId(self,visualization_id:int):
        x = self.gc["challenges"]
        result = x[x['visualizations'] == visualization_id]
        return result

    def getVisualizationById(self,id):
        return self.gc["visualizations"][self.gc["visualizations"]['id'] == id].iloc[0]
    
    def getVisualizationChallenges(self,visualization):
        return self.getVisualizationChallengesbyId(visualization['id'])

    def getVisualizationInitialChallenges(self,visualization):
        result = [c for _,c in self.getVisualizationChallenges(visualization).iterrows() if self.isChallengeInitialLevel(c)]
        return result
    
    def getVisualizationTerminalChallenges(self,visualization):
        return [c for _, c in self.getVisualizationChallenges(visualization).iterrows() if self.isChallengeTerminalLevel(c)]

    def challengeEqual(self,c1,c2):
        return (c1['id'] == c2['id'])

    def reachableChallenges(self,challenge):
        result,visited = self.reachableChallengesIntern(challenge,visitedids=[])
        return result
    
    
    def reachableChallengesIntern(self,challenge,visitedids=[]):
        if (challenge['id'] in visitedids):
            result = []
        elif (self.isChallengeTerminalLevel(challenge)): 
            result = [challenge]
            visitedids = [*visitedids, challenge['id']]
        else:
            result = []
            visitedids = [*visitedids, challenge['id']]
            nexts = [self.getChallengeSuccessChallenge(challenge),self.getChallengeFailureChallenge(challenge)]
            nexts = [n for n in nexts if (n is not None or next['id'] not in visitedids)]
            for c in nexts:
                res,visitedids = self.reachableChallengesIntern(c,visitedids)
                result.extend(res)
        
        return result,visitedids
        
    def reachable(self, fromchallenge,tochallenge,successonly=True,visitedids=[]):
        
        if (self.challengeEqual(fromchallenge,tochallenge)):
            result = True
        else:
            nexts = []
            if (not self.isChallengeTerminalLevel(fromchallenge)):
                nexts = [self.getChallengeSuccessChallenge(fromchallenge)]
            if (not successonly):
                nexts.append(self.getChallengeFailureChallenge(fromchallenge))
            nexts = [n for n in nexts if (n is not None or next['id'] not in visitedids)]
            visitedids = [*visitedids, fromchallenge['id']]
            result = False
            for c in nexts:
                if (self.reachable(c,tochallenge,successonly=successonly,visitedids=visitedids)):
                    result = True
                    break
        
        return result
    
    def initialChallengeReachedTerminalChallenges(self,visualization,challenge,successonly=True):
                
        intern = []
        
        for t in self.getVisualizationTerminalChallenges(visualization):
            if (self.reachable(challenge,t,successonly=successonly,visitedids=[])):
                intern.append(t)
        
        return intern

            
    def allInitialChallengesNotReachingTerminalChallenge(self,visualization,successonly=True):
        return [i for i in self.getVisualizationInitialChallenges(visualization) if self.initialChallengeReachedTerminalChallenges(visualization,i,successonly=successonly) == []]
    
    def terminalChallengeReachedFromInitialChallenges(self,visualization,challenge,successonly=True):
        intern = []
        for i in self.getVisualizationInitialChallenges(visualization):
            if (self.reachable(i,challenge,successonly=successonly,visitedids=[])):
                intern.append(i)
        return intern

            
    def allTerminalChallengesNotReachedFromInitialChallenge(self, visualization,successonly=True):
        return [t for t in self.getVisualizationTerminalChallenges(visualization) if self.terminalChallengeReachedFromInitialChallenges(visualization,t,successonly=successonly) == []]
                                
    """"
    ====================================================

    Challenges

    ====================================================
    """

    def getChallenges(self):
        return self.gc["challenges"]

    def getChallengeWithId(self,challenge_id):
        challenges = self.gc["challenges"]
        result = challenges[challenges['id'] == challenge_id]
        return result.iloc[0]

    def getChallengeTasks(self,challenge_id):
        tasks = self.gc["tasks"]
        return tasks[tasks['challenge'] == challenge_id]

    def isChallengeInitialLevel(self,challenge):
        return challenge['is_initial_level'] == 1

    def getChallengeSuccessChallenge(self,challenge):
        return self.getChallengeWithId(challenge['success_next'])
        
    def getChallengeFailureChallenge(self,challenge):
        return self.getChallengeWithId(challenge['failure_next'])

    def getChallengeTargetPoints(self,challenge):
        return challenge['target']

    def getChallengeDuration(self,challenge):
        return challenge['evaluate_fail_every_x_minutes']

    def isChallengeTerminalLevel(self,challenge):
        return self.challengeEqual(challenge,self.getChallengeSuccessChallenge(challenge))

    def computeChallengeReachablePoints (self,challenge):
        if (not math.isnan(self.getChallengeDuration(challenge))):
            days_for_level = self.getChallengeDuration(challenge) / (24 * 60)
            result = 0
            for _,t in self.getChallengeTasks(challenge['id']).iterrows():
                points = self.computeTaskMaximumAchievablePoints(t,days_for_level)
                if (points is not None):
                    result = result + points
                else:
                    result = None
                    break
            return result
        else:
            return None
    
    def getChallengeEditURL(self,visualization,challenge):
        return f"https://campaigns.healthyw8.gamebus.eu/editor/for/{visualization['campaign']}/{challenge['visualizations']}/challenges/{challenge['id']}"
    
    """"
    ====================================================

    Tasks

    ====================================================
    """

    def getTasks(self):
        return self.gc["tasks"]
    
    def getTask_points(self,task):
        return task["points"]

    def getTask_max_times_fired(self,task):
        return task["max_times_fired"]

    def getTask_min_days_between_fire(self,task):
        return task["min_days_between_fire"]

    def computeTaskMaximumAchievablePoints(self,task,days_for_level):
        reset_of_reward_counter = self.getTask_max_times_fired(task) # days
        reward_count = self.getTask_max_times_fired(task)
        time_window_reseting_award_count = self.getTask_min_days_between_fire(task)
        if (not math.isnan(days_for_level) and 
            not math.isnan(reset_of_reward_counter) and 
            not math.isnan(reward_count)):
            p = math.floor(days_for_level / time_window_reseting_award_count)
            max_number_of_times_for_task = p*reward_count + min(days_for_level-(p*time_window_reseting_award_count),reward_count)
            max_points_for_task = max_number_of_times_for_task * self.getTask_points(task)
            return max_points_for_task
        else:
            return None

    """"
    ====================================================

    checks

    ====================================================
    """
    
    """
    spellchecking names and descriptions of tasks and challenges
    """
    
    knownwords = [
        "Newbie","Rookie","Expert","Master","Grandmaster","Skilled","Proficient",
        "Mini-Walks","Unterarmstützer","Plank","Wall-Sit","Joggen","MIND","kalziumreiche","Freundlichkeits-Tagebuch","Mikrobiota","Steppin","Up","Joggingintervalle","Wandsitzer","Wandsitz"   
    ]
    
    def spellcheckTaskAndChallenges(self):
        logging.info("Spellchecking tasks and challenges")
        tool = language_tool_python.LanguageTool(language='de-DE',mother_tongue="de-DE",new_spellings=self.knownwords,remote_server='http://localhost:8081/')
        for _,t in self.getTasks().iterrows():
            c = self.getChallengeWithId(t['challenge'])
            vis = self.getVisualizationById(c['visualizations'])
            error,errormessage = self.checkText(tool,t['name'],'Name of task')
            if (error):
                self.addError(SPELLCHECKER,f"{errormessage}",vis,c)
            # error,errormessage = self.checkText(tool,t['description'],'Description of task')
            # if (error):
            #     self.addError(SPELLCHECKER,f"Description of {errormessage}",vis,c)
        for _,c in self.getChallenges().iterrows():
            vis = self.getVisualizationById(c['visualizations'])
            error,errormessage = self.checkText(tool,c['name'],'Name of challenge')
            if (error):
                self.addError(SPELLCHECKER,f"{errormessage}",vis,c)
            # error,errormessage = self.checkText(tool,c['description'],'Description of challenge')
            # if (error):
            #     self.addError(SPELLCHECKER,f"Description of {errormessage}",vis,c)
        logging.info(self.checkResult(SPELLCHECKER))
        
    def checkText(self,tool,text,type):
        if isinstance(text, float):
            errormessage = f"{type} is empty."
        else:
            checkmatches = tool.check(text)
            testname = classify_matches(checkmatches)
            errormessage = None
            if (testname==TextStatus.FAULTY):
                correction = tool.correct(text)
                errormessage = f"{type} is faulty '{text}'. Proposed correction is \n'{correction}'"
            elif (testname==TextStatus.GARBAGE):
                errormessage = f"{type} is garbage '{text}', no correction can be proposed"
        return (errormessage is not None), errormessage

    """
    check that all tasks that have Gamebus Studio as data provider have a secret and if two tasks have the same secret then they are the same in terms of description. 
    """
    
    def splittriple(self,triple):
        pt = triple.split(',')
        if (len(pt)>3):
            res = [pt[0], pt[1], ",".join(pt[2:])]
            # print(f"Fixing {triple} to {res}")
            return res
        else:
            return pt
        
    def parseConditionsIntoTriples(self,str):
        # Find all triples
        if isinstance(str, float):
            return []
        else:
            triples = re.findall(r'\[([^\]]+)\]', str)
            # Split each triple into its components
            parsed_triples = [self.splittriple(triple) for triple in triples]
            return parsed_triples
        
    def conditionTriplesFindSecret(self,triples):
        
        secret = None
        for triple in triples:
            if (triple[0]=="SECRET" and triple[1]==" EQUAL"):
                secret = triple[2]
                break
        return secret 
    
    def checkTasksHaveSecrets(self,fixemptysecrets):
        logging.info("Checking tasks secrets")
        gamebusstudio = "GameBus Studio"
        secretchecks = {}
        for row,t in self.getTasks().iterrows():
            secret = self.conditionTriplesFindSecret(self.parseConditionsIntoTriples(t['conditions']))
            if (t['dataproviders']==gamebusstudio):
                if (secret is not None):
                    if (secret in secretchecks):
                        secretchecks[secret].append(t)
                    else:
                        secretchecks[secret] = [t]
                else:
                    c = self.getChallengeWithId(t['challenge'])
                    vis = self.getVisualizationById(c['visualizations'])
                    news = str(t['name']).replace(" ","-").replace("ü","ue").replace("ä","ae").replace("ö","oe").replace("ß","sz").replace(".","-dot-").replace(";","-semicolon-").replace(":","-colon-")
                    proposedsecret = f"[SECRET, EQUAL, {news}]"
                    if (isinstance(t['conditions'],str)):
                        proposedsecret = f"{proposedsecret}, {t['conditions']}"
                    if (fixemptysecrets):
                        self.gc["tasks"].at[row,'conditions'] = proposedsecret
                    self.addError(SECRETS,f"Task '{t['name']}' has no secret. Proposing {proposedsecret} at column 'conditions' in row={row} (name={self.gc["tasks"].at[row,'name']})",vis,c)
                    
        for key in secretchecks.keys():
            challenges = [secretchecks[key][0]['challenge']]
            if (len(secretchecks[key])>1):
                res = True
                for t in secretchecks[key][1:]:
                    challenges.append(t['challenge'])
                    res = res and (t['name']==secretchecks[key][0]['name'])     
                if (not res):
                    t = secretchecks[key][0]
                    c = self.getChallengeWithId(t['challenge'])
                    vis = self.getVisualizationById(c['visualizations'])
                    print(challenges)
                    p = [f"{t} ({self.getChallengeWithId(t)['name']})" for t in challenges]
                    self.addError(SECRETS,f"Task '{t['name']}' has copies with the same secret '{key}', but that have different names (see challenges {p})",vis,c)               
        logging.info(self.checkResult(SECRETS))

    """
    reachability check: Checking that in all visualization every initial level reaches a terminal level and every 
    terminal level can be reached from some initial level
    """

    def checkInitialAndTerminalReachability(self):
        logging.info("Checking reachability of initial and terminal challenges")
        for _, vis in self.getVisualizations().iterrows(): 
            for c in self.allInitialChallengesNotReachingTerminalChallenge(vis):
                self.addErrors(REACHABILITY,"Initial Challenge without terminal challenge",vis,c)
            for c in self.allTerminalChallengesNotReachedFromInitialChallenge(vis):
                self.addError(REACHABILITY,"Terminal Challenge not reachable from any initial challenge",vis,c)
        logging.info(self.checkResult(REACHABILITY))
    
    """
    visualizationintern check: Checking that for all visualizations, all challenges reachable from an intial level all belong to the same visualization and label.
    """
    
    def checkAllReachableChallengesAreInSameVisualizationAndLabel(self):
        logging.info("Checking reachable challenges are in same visualization and label")      
        for _, vis in self.getVisualizations().iterrows(): 
            for c in self.getVisualizationInitialChallenges(vis):
                for reachable in self.reachableChallenges(c):
                    if (c['visualizations'] == reachable['visualizations'] and ((isinstance(c['labels'],float) and isinstance(reachable['labels'],float)) or c['labels'] == reachable['labels'])):
                        break
                    else:
                        description = f"Initial challenge visualization = '{c['visualizations']}'; reachable challenge visualization = '{reachable['visualizations']}'{os.linesep} Initial challenge labels = '{c['labels']}'; reachable challenge labels = '{reachable['labels']}'{os.linesep}"
                        self.addError(VISUALIZATIONINTERN,"Reachable Challenge from some initial level is not in same visualization or not with same label:\n"+description,vis,reachable)
        logging.info(self.checkResult(VISUALIZATIONINTERN))
                
                
    """
    consistency check: Checking that all initial levels have failure successor to themselves and all terminal levels have successlevel to themselves
    """
    
    def checkIntialandTerminalLevelConsistentSuccessors(self):
        logging.info("Checking self-reference for initial levels are also failure successor and terminal levels are also successlevel")       
        for _, vis in self.getVisualizations().iterrows(): 
            for c in self.getVisualizationInitialChallenges(vis):
                if (not self.challengeEqual(c,self.getChallengeFailureChallenge(c))):
                    self.addErrors(CONSISTENCY,
                                f"Initial challenge does not lead to itself on failure {self.getChallengeFailureChallenge(c)['id']}",vis,c)
            for c in self.getVisualizationTerminalChallenges(vis):
                if (not self.challengeEqual(c,self.getChallengeSuccessChallenge(c))):
                    self.addError(CONSISTENCY,f"Terminal challenge does not lead to itself on success {self.getChallengeSuccessChallenge(c)['id']}",vis,c)
        logging.info(self.checkResult(CONSISTENCY)) 
    
    """
    challenge success transition check: checking that the way the tasks in a level are configured (points, reward count, reset time for reward count) all together allow to achieve the target points defined for the level.
    
    """
    
    def checkChallengeTargetPointsCanBeReached(self):
        logging.info("Checking that target points of levels can be reached with tasks")
        for _, vis in self.getVisualizations().iterrows(): 
            for _, c in self.getVisualizationChallenges(vis).iterrows(): 
                challenge_target_points = self.getChallengeTargetPoints(c)
                challenge_tasks_reachable_points = self.computeChallengeReachablePoints(c)
                if (not math.isnan(challenge_target_points)):
                    if (challenge_tasks_reachable_points is not None):
                        if (challenge_tasks_reachable_points < challenge_target_points):
                            self.addError(TARGETPOINTSREACHABLE,
                                        f"Challenge target points ({challenge_target_points}) cannot be reached with tasks (max reachable is {challenge_tasks_reachable_points})",
                                        vis,c)
                    else:
                        self.addError(TARGETPOINTSREACHABLE,
                                    f"Challenge reachable points ({challenge_tasks_reachable_points}) cannot be computed, missing values in tasks.",vis,c)
                else:
                    self.addError(TARGETPOINTSREACHABLE,
                                f"Challenge no target points defined ({challenge_target_points}).",vis,c)
        logging.info(self.checkResult(TARGETPOINTSREACHABLE))    

    """
    check TTM structure: checking that the way the tasks in a challenge follow the TTM structure
    
    """

    def checkTTMstructure(self,norelapselevels=4):
        assert(norelapselevels>0)
        logging.info("Checking all visualizations are in TTM structure")
        for _, vis in self.getVisualizations().iterrows(): 
                for c in self.getVisualizationInitialChallenges(vis):
                    self.checkchallengeTTM(vis,c,norelapselevels)
        logging.info(self.checkResult(TTMSTRUCTURE))  
                    
    def checkchallengeTTM(self,vis,c,norelapselevels=0,lastlevel=None):
        nextlevel = self.getChallengeSuccessChallenge(c)
        if (norelapselevels>0):
            # Normal level, check failure level is itself and continue with success level
            if (not self.challengeEqual(self.getChallengeFailureChallenge(c),c)): # this is only for non final levels. 
                self.addError(TTMSTRUCTURE, f"Challenge {c['id']} ({c['name']}) should have failure success level to itself.",vis,c)
            elif (not self.challengeEqual(nextlevel,c)):
                self.checkchallengeTTM(vis,nextlevel,norelapselevels-1,c)
            else: 
                return
        elif (self.challengeEqual(nextlevel,c)):
            if (not self.challengeEqual(self.getChallengeFailureChallenge(c),lastlevel)): # this is only for final levels. 
                self.addError(TTMSTRUCTURE, f"Challenge {c['id']} ({c['name']}) should have failure to previous level {lastlevel['id']} ({lastlevel['name']}) that led to {c['id']} ({c['name']}) as successor level.",vis,c)
        else:
            # Relapse Level, check
            # Challengefailurelevel is a relapse level (not lastlevel)
            relapselevel = self.getChallengeFailureChallenge(c)
            relapselevelfailure = self.getChallengeFailureChallenge(relapselevel)
            relapselevelsuccess = self.getChallengeSuccessChallenge(relapselevel)
            
            if (self.challengeEqual(relapselevel,lastlevel)): 
                self.addError(TTMSTRUCTURE,f"Challenge {c['id']} ({c['name']}): its 'At risk level' {relapselevel['id']} ({relapselevel['name']}) should not be the previous level {lastlevel['id']} ({lastlevel['name']}) in TTM hierarchy. Maybe the error is also from that successor challenge being wrong.",vis,c)
            if (not self.challengeEqual(relapselevelfailure,lastlevel)):
                self.addError(TTMSTRUCTURE,f"Challenge {c['id']} ({c['name']}): its 'At risk level' {relapselevel['id']} ({relapselevel['name']}) should have as failure challenge the previous level {lastlevel['id']} ({lastlevel['name']}) in the TTM hierarchy that led to {c['id']} ({c['name']}).",vis,c)
            if (not self.challengeEqual(relapselevelsuccess,c)):
                self.addError(TTMSTRUCTURE,f"Challenge {c['id']} ({c['name']}): its 'At risk level' {relapselevel['id']} ({relapselevel['name']}) should have as success challenge the challenge {c['id']} ({c['name']})  again.",vis,c)
            self.checkchallengeTTM(vis,nextlevel,0,c)
                

def main(filename, consistency, visualizationintern, reachability, 
        targetpointsreachable, secrets, fixsecrets:bool, spellchecker:bool, ttm:bool):
    
    logging.info("Version 0.2")
    gc = CampaignChecker(filename)
    fixsecrets = False
    spellchecker = False
    all = not (consistency or visualizationintern or reachability or 
            targetpointsreachable or secrets or spellchecker or ttm) 
    if (reachability or all):
        gc.checkInitialAndTerminalReachability()
    if (consistency or all):
        gc.checkIntialandTerminalLevelConsistentSuccessors()
    if (visualizationintern or all):
        gc.checkAllReachableChallengesAreInSameVisualizationAndLabel()
    if (targetpointsreachable or all):
        gc.checkChallengeTargetPointsCanBeReached()
    if (secrets or all):
        gc.checkTasksHaveSecrets(fixsecrets)
    if (spellchecker): # making spellchecker to be explicitely triggered as it is very slow.
        gc.spellcheckTaskAndChallenges()
    if (ttm or all): 
        gc.checkTTMstructure()
    gc.errorsToLog()
    gc.errorsToExcel()
    if (fixsecrets):
        gc.campaignToExcel()
    
if __name__ == "__main__":

    if os.path.isfile(LOG_NAME):
        os.remove(LOG_NAME)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_NAME), # Write to a log file
            logging.StreamHandler() # Also write to stdout as usual
        ]
    )
    
    argparser = argparse.ArgumentParser(
        prog="Gamebus Campaign Checker",
        description="Python Script to verify properties of GameBus Campaigns."
    )
        
    argparser.add_argument("-c",f"--{CONSISTENCY}",action='store_true',
                            help="Checking that all initial levels have failure successor to themselves and all terminal levels have successlevel to themselves")

    argparser.add_argument("-v",f"--{VISUALIZATIONINTERN}",action='store_true',
                            help="Checking that for all visualizations, all challenges reachable from an initial level all belong to the same visualization and label.") 
    
    argparser.add_argument("-r",f"--{REACHABILITY}",
                            action='store_true',
                            help="Checking that in all visualization every initial level reaches a terminal level and every terminal level can be reached from some initial level")

    argparser.add_argument("-t",f"--{TARGETPOINTSREACHABLE}",action='store_true',
                            help="Checking that the way the tasks in a level are configured (points, reward count, reset time for reward count) all together allow to achieve the target points defined for the level.")

    argparser.add_argument("-s",f"--{SECRETS}",action='store_true',
                            help="Checking that the all tasks provided by gamebus Studio have secrets")
    
    argparser.add_argument("-ttm",f"--{TTMSTRUCTURE}",action='store_true',
                            help="Checking that the all visualizations have the TTM structure")
    
    argparser.add_argument("-fs",f"--fix{SECRETS}",action='store_true',
                            help="Fixes empty Secrets and save into campaign file.")
    
    argparser.add_argument("-g",f"--{SPELLCHECKER}",action='store_true',
                            help="Spellchecking tasks and challenges names and descriptions")
    
    argparser.add_argument("-f", "--filename",
                            help="The path to the EXCEL file containing the GameBus Campaign Description.")
    
    args = argparser.parse_args()
    if (args.filename is None): 
        argparser.print_usage()
    else: 
        main(args.filename, args.consistency, args.visualizationintern, args.reachability, 
            args.targetpointsreachable,args.secrets,args.fixsecrets,args.spellchecker,args.ttm)

