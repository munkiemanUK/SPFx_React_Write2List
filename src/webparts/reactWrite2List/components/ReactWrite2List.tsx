import * as React from 'react';
import styles from './ReactWrite2List.module.scss';
import { IReactWrite2ListProps } from './IReactWrite2ListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { taxonomy, ITerm, ITermSet, ITermStore } from '@pnp/sp-taxonomy';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

export interface IPTerm {
  parent?: string;
  id: string;
  name: string;
}

export interface IReactTaxonomyState {
  terms: IPTerm[];
  tags: IPickerTerms;
}

export default class ReactWrite2List extends React.Component<IReactWrite2ListProps, IReactTaxonomyState, {}> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      terms: [],
      tags: []
    };
  }

  public async onDivisionTaxPickerChange(terms: IPickerTerms): Promise<any> {
    const data = {};
    //let termLabel:string="";

    data['Tags'] = {
      "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
      "Label": terms[0].name,
      'TermGuid': terms[0].key,
      'WssId': '-1'
    };
    //alert('updated Divsion with: '+terms[0].name+' WPI ID='+this.state.ItemID);

    const item: any = await sp.web.lists.getByTitle("WPI_TopSection").items.getById(2).get();
    alert("title="+item.Title);
    //termLabel=terms[0].name;
    if(item.Title == ""){
      alert('data written ' + terms[0].key);
      return (await sp.web.lists.getByTitle("WPI_TopSection").items.add({Title: terms[0].name}));
    } else {
      alert('data updated ' + terms[0].key);
      return await sp.web.lists.getByTitle("WPI_TopSection").items.getById(2).update({WPI_ID: 6000,Title: terms[0].name});
    }

    //alert('data written to ID:6152 ' + terms[0].key);
    //termLabel=terms[0].name;

    //return await sp.web.lists.getByTitle("HandS_WPI_Forms").items.getById(6152).update({Division: terms[0].name});
    //return (await sp.web.lists.getByTitle("Test_Metadata").items.add({Title: termLabel}))
  }  

  public render(): React.ReactElement<IReactWrite2ListProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    
    return (
      <section className={`${styles.reactWrite2List} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <TaxonomyPicker allowMultipleSelections={false}
            initialValues={this.state.tags}
            termsetNameOrID="Divisions"
            panelTitle="Select Division"
            label="Division"
            context={this.props.context}
            onChange={this.onDivisionTaxPickerChange}
            isTermSetSelectable={false} />
        </div>  
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
