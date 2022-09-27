import * as React from 'react';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TooltipHost, ITooltipHostStyles } from 'office-ui-fabric-react/lib/Tooltip';
import { getTheme, mergeStyleSets } from 'office-ui-fabric-react';
import styles from '.././safetyHub/components/SafetyHub.module.scss';
export interface IToolTipPrrops {
    value?: string;
    lenght?: number;
}

const calloutProps = { gapSpace: 0 };
// The TooltipHost root uses display: inline by default.
// If that's causing sizing issues or tooltip positioning issues, try overriding to inline-block.
const theme = getTheme();
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
const classNames = mergeStyleSets({
    // Applied to make content overflow (and tooltips activate)
    overflow: {
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'normal',
        width: 330,
    },


});

export default class Tooltip extends React.Component<IToolTipPrrops, {}>
{
    public constructor(props: IToolTipPrrops, state: {}) {
        super(props);
    }

    public render(): React.ReactElement<IToolTipPrrops> {
        return (
            <div>
                <TooltipHost
                    content={this.props.value}
                    calloutProps={calloutProps}
                    styles={hostStyles}
                    hostClassName={classNames.overflow}
                >
                    <div className="tooltip">
                        {this.props.value && this.props.value.length > this.props.lenght ?
                            this.props.value.substring(0, this.props.lenght) + "..."
                            : this.props.value

                        }
                    </div>
                </TooltipHost>
            </div>

        );
    }
}
