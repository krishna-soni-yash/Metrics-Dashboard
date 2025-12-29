import * as React from 'react';
import styles from './MetricsDashboard.module.scss';
import type { IMetricsDashboardProps } from './IMetricsDashboardProps';
import Dashboard from './Dashboard/Dashboard'; // added import
//import Home from './Home';


export default class MetricsDashboard extends React.Component<IMetricsDashboardProps> {
  public render(): React.ReactElement<IMetricsDashboardProps> {
    const {
    context
    } = this.props;

    return (
      <section className={`${styles.metricsDashboard}`}>
        {/* Render the Dashboard component */}
        <Dashboard context={context} />
        {/* <Home></Home> */}
      </section>
    );
  }
}
