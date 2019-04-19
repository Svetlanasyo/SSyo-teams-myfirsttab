import * as React from 'react';
import styles from './MyFirstTeamsTab.module.scss';
import { IMyFirstTeamsTabProps } from './IMyFirstTeamsTabProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jQuery from 'jquery';
import 'jqueryui';
import './../../../../styles/jquery-ui-1.11.4.custom/jquery-ui.css';
import { TextField } from 'office-ui-fabric-react/lib';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface IMyFirstTeamsTabState {
  firstNumber?: number;
  secondNumber?: number;
}

export default class MyFirstTeamsTab extends React.Component<IMyFirstTeamsTabProps, IMyFirstTeamsTabState, {}> {
  
  public componentDidMount() {
            
    this.setState({});
  

      const accordionOptions: JQueryUI.AccordionOptions = {
          animate: true,
          collapsible: false,
          icons: {
            header: 'ui-icon-circle-arrow-e',
            activeHeader: 'ui-icon-circle-arrow-s'
          }
        };
      
        jQuery('.accordion').accordion(accordionOptions);
  }

  public render(): React.ReactElement<IMyFirstTeamsTabProps> {
    
        let result: string[] = this.props.resultStack.slice(-3);
    
        return (
          <div className={ styles.myFirstTeamsTab }>
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                  <div className={styles.accordion + " accordion"}>
                    <h3>{escape(this.props.description)}</h3>
                    <div>
                      <TextField 
                        label="Pierwsza liczba"
                        type="number"
                        onChanged={value => this.setState({firstNumber: Number(value)})}
                        placeholder="Wprowadż liczbę"
                      />
                      <TextField
                       label="Druga liczba"
                       type="number"
                       onChanged={value => this.setState({secondNumber: Number(value)})}
                       placeholder="Wprowadż liczbę"
                      />
                      <br/>
                      <div>
                        <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="+"
                          onClick={() => this.calculate((first, second) => first+second)}
                        />
                          <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="-"
                          onClick={() => this.calculate((first, second) => first-second)}
                        />
                          <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="/"
                          onClick={() => this.divide((first, second) => first/second)}
                        />
                          <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="*"
                          onClick={() => this.calculate((first, second) => first*second)}
                        />
                          <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="%"
                          onClick={() => this.divide((first, second) => first%second)}
                        />
                          <PrimaryButton
                          disabled={false}
                          checked={false}
                          text="PI"
                          onClick={() => this.calculate((first, second) => 3.14)}
                        />
                      </div>
                    </div>
                    <h3>Wynik 1</h3>
                    <div>
                        <p>
                            {result[2]}
                        </p>
                    </div>
                    <h3>Wynik 2</h3>
                    <div>
                        <p>
                            {result[1]}
                        </p>
                    </div>
                    <h3>Wynik 3</h3>
                    <div>
                        <p>
                            {result[0]}
                        </p>
                    </div>
                </div>
                </div>
              </div>
            </div>
          </div>
        );
      }
      
      private isInputNotEmpty(): boolean {
        let first: number = this.state.firstNumber;
        let second: number = this.state.secondNumber;
        if (isNaN(first) || isNaN(second)) {
          return false;
        }
        return true;
      }
    
      private calculate(operation: (num1: number, num2: number) => number) {
        if (this.isInputNotEmpty()) {
          let first: number = this.state.firstNumber;
          let second: number = this.state.secondNumber;
          let result = operation(first, second)
          this.props.resultStack.push(result+'');
          this.forceUpdate();
        } else {
        alert('Wszystkie pola powinny być uzupełnione');
        }
      }
    
    
    private divide(operation: (num1: number, num2: number) => number) {
      let first: number = this.state.firstNumber;
      let second: number = this.state.secondNumber;
      if (second == 0) {
        alert('Wartość drugiej liczby nie może być równa 0')
      } else if (this.isInputNotEmpty()) {
        let result = operation(first, second)
        this.props.resultStack.push(result+'');
        this.forceUpdate();
      } else {
      alert('Wszystkie pola powinny być uzupełnione');
      }
    }
  }


    
  

