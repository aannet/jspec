# JSPEC

## JSPECS READER

The goal of this app is to generate dynamic specification from an excel spreadsheet containing **user stories** and **business rules**.

### TODO
- [ ] Load XSLX file from an app window and not from code
- [ ] Add search (jqmark ? lodash ?) 
- [ ] Create multistep form to configure mapping before displaying 

### Actual launch
```
npm run start
```


### Actual Data Structure

```
const dataUsStatic = {
    labels: {
        usAsLabel: "En tant que2",
        usToLabel: "Afin de",
        usICanLabel: "Je peux",
        usCommentsLabel: "Commentaire"
    },
    usList: [
        {
            id: "ID_US_1",
            usAs: "developpeur",
            usTo: "de devenir lead",
            usIcan: "faire de la veille",
            usComments: "Mes commentaires",
            rmList: [
                {
                    id: "RM_01",
                    rmText: "Ma règle 1"
                },
                {
                    id: "RM_02",
                    rmText: "Ma règle 2"
                }
            ]
        },
        {
            id: "ID_US_2",
            usAs: "product owner",
            usTo: "de travailler efficacement avec l'équipe",
            usIcan: "être disponible",
            usComments: "Mes commentaires2",
            rmList: [
                {
                    id: "RM_03",
                    rmText: "Ma règle 3"
                }
            ]
        }
    ]
};

const dataRMStatic = {
    rmList: [
        {
            id: "RM_01",
            rmText: "Ma règle 1",
            idUS: "ID_US_1"
        },
        {
            id: "RM_02",
            rmText: "Ma règle 2",
            idUS: "ID_US_1"
        },
        {
            id: "RM_03",
            rmText: "Ma règle 3",
            idUS: "ID_US_2"
        },
    ]
};
```
