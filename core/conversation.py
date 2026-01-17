from typing import List, Dict, Optional
from datetime import datetime

class ConversationManager:
    def __init__(self, system_prompt: Optional[str] = None, max_history_length: int = 20):
        self.messages: List[Dict] = []
        self.created_at = datetime.now()
        self.max_history_length = max_history_length

        if system_prompt:
            self.add_system_message(system_prompt)
        
    def add_system_message(self,content:str):
        message={
            'role':'system',
            'content':content
        }
        if self.messages and self.messages[0]['role']=='system':
            self.messages[0]=message
        else:
            self.messages.insert(0,message)
    
    def add_user_message(self,content:str):

        self.messages.append({
            'role':'user',
            'content':content,
            'timestamp':datetime.now().isoformat()
        })
        self._trim_history()
    def add_assistant_message(self,content:str):
        self.messages.append({
            'role':'assistant',
            'content':content,
            'timestamp':datetime.now().isoformat()
        })
        self._trim_history()
    def _trim_history(self):
        system_msgs=[m for m in self.messages if m['role']=='system']
        other_msgs=[m for m in self.messages if m['role']!='system']

        if len(other_msgs) > self.max_history_length:
            other_msgs = other_msgs[-self.max_history_length:]
        self.messages = system_msgs + other_msgs
        
    def get_messages(self)->List[Dict]:
        return [
            {'role':msg['role'],'content':msg['content']} for msg in self.messages
        ]
    def get_conversation_length(self)->int:
        return len([m for m in self.messages if m['role'] != 'system'])
    
    def clear_conversation(self,keep_system:bool=True):
        if keep_system and self.messages and self.messages[0]['role']=='system':
            system_msg =self.messages[0]
            self.messages=[system_msg]
        else:
            self.messages=[]
    def  get_last_n_messages(self,n:int)->List[Dict]:
        system_msg=[m for m in self.messages if m['role']=='system']
        other_msgs=[m for m in self.messages if m['role']!='system']
        last_n_msgs=other_msgs[-n:]
        return system_msg + last_n_msgs
    
    def print_conversation(self):
        print("Conversation History:")

        for i,msg in enumerate(self.messages,1):
            role={'system':'[System]','user':'[User]','assistant':'[Assistant]'}.get(msg['role'],'[Unknown]')
            content = msg['content'][:100] + "..." if len(msg['content']) > 100 else msg['content']
            print(f"{i}. {role}: {content}")
        print(f"總共{self.get_conversation_length()}輪對話")





